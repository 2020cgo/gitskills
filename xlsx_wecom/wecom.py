"""把均值摘要和 xlsx 文件发给企业微信里的自己。"""

from __future__ import annotations

import json
import os
import uuid
from dataclasses import dataclass
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit, urlunsplit
from urllib.request import Request, urlopen

from xlsx_wecom.means import Outcome, format_means

API = "https://qyapi.weixin.qq.com"
_TEXT_LIMIT = 2048
_HINTS = {
    60020: "企业微信不允许当前出口 IP 调用接口。到应用的「企业可信 IP」里加上这台机器的公网 IP。",
    81013: "接收人不在该应用的可见范围内。把你的账号加进应用可见范围。",
    60011: "应用没有权限发给这个账号。检查可见范围，并确认 WECOM_USER_ID 是你自己的账号。",
    93000: "机器人 webhook 地址无效，或 key 已失效。",
}


class WeComNotConfigured(Exception):
    """还没有企业微信凭证。"""


class WeComError(Exception):
    """企业微信接口拒绝了这次发送。"""


@dataclass(frozen=True)
class AppCredentials:
    corp_id: str
    secret: str
    agent_id: int
    user_id: str


@dataclass(frozen=True)
class WebhookCredentials:
    url: str


def load_local_env(path: Path) -> None:
    """读取 .env。已存在的环境变量优先，不覆盖。"""
    if not path.is_file():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and value:
            os.environ.setdefault(key, value)


def resolve_credentials() -> AppCredentials | WebhookCredentials:
    corp_id = os.environ.get("WECOM_CORP_ID", "").strip()
    secret = os.environ.get("WECOM_SECRET", "").strip()
    agent_raw = os.environ.get("WECOM_AGENT_ID", "").strip()
    user_id = os.environ.get("WECOM_USER_ID", "").strip()
    webhook = os.environ.get("WECOM_WEBHOOK", "").strip()
    app_parts = [corp_id, secret, agent_raw, user_id]
    if any(app_parts):
        if not all(app_parts):
            raise WeComNotConfigured(
                "发给自己需要同时设置 WECOM_CORP_ID、WECOM_SECRET、WECOM_AGENT_ID、WECOM_USER_ID。"
                "现在只设置了其中一部分。"
            )
        if user_id == "@all":
            raise WeComError("WECOM_USER_ID 不能是 @all。当前需求是发给你自己，请填写你的账号。")
        if not agent_raw.isdigit():
            raise WeComError(f"WECOM_AGENT_ID 必须是数字，现在是「{agent_raw}」。")
        return AppCredentials(corp_id, secret, int(agent_raw), user_id)
    if webhook:
        _validate_webhook(webhook)
        return WebhookCredentials(webhook)
    raise WeComNotConfigured(
        "企业微信还没配置，文件没有发送。\n"
        "发给自己：设置 WECOM_CORP_ID、WECOM_SECRET、WECOM_AGENT_ID、WECOM_USER_ID。\n"
        "或暂时发到群：设置 WECOM_WEBHOOK。"
    )


def send_outcome(
    outcome: Outcome,
    credentials: AppCredentials | WebhookCredentials,
    opener=urlopen,
) -> str:
    if outcome.output is None:
        raise WeComError("没有生成文件，未发送。")
    summary = _clip(format_means(outcome))
    filename = outcome.output.name
    payload = outcome.output.read_bytes()
    if isinstance(credentials, AppCredentials):
        token = _get_token(credentials, opener)
        media_id = _upload_app(token, filename, payload, opener)
        _post_json(
            f"{API}/cgi-bin/message/send?access_token={token}",
            {
                "touser": credentials.user_id,
                "msgtype": "text",
                "agentid": credentials.agent_id,
                "text": {"content": summary},
            },
            opener,
        )
        _post_json(
            f"{API}/cgi-bin/message/send?access_token={token}",
            {
                "touser": credentials.user_id,
                "msgtype": "file",
                "agentid": credentials.agent_id,
                "file": {"media_id": media_id},
            },
            opener,
        )
        return f"已通过企业微信发给 {credentials.user_id}。"
    media_id = _upload_webhook(credentials.url, filename, payload, opener)
    _post_json(credentials.url, {"msgtype": "text", "text": {"content": summary}}, opener)
    _post_json(credentials.url, {"msgtype": "file", "file": {"media_id": media_id}}, opener)
    return "已发到企业微信群机器人。这是群消息，不是单发给你自己。"


def _validate_webhook(url: str) -> None:
    parts = urlsplit(url)
    if parts.scheme != "https" or parts.netloc != "qyapi.weixin.qq.com":
        raise WeComError("WECOM_WEBHOOK 必须是 https://qyapi.weixin.qq.com 的企业微信机器人地址。")
    if not parts.path.endswith("/cgi-bin/webhook/send"):
        raise WeComError("WECOM_WEBHOOK 应该是机器人的 send 地址。")


def _get_token(credentials: AppCredentials, opener) -> str:
    url = (
        f"{API}/cgi-bin/gettoken?corpid={credentials.corp_id}"
        f"&corpsecret={credentials.secret}"
    )
    payload = _read_json(url, opener, method="GET")
    token = payload.get("access_token")
    if not token:
        raise WeComError("企业微信没有返回 access_token。")
    return str(token)


def _upload_app(token: str, filename: str, payload: bytes, opener) -> str:
    url = f"{API}/cgi-bin/media/upload?access_token={token}&type=file"
    body = _read_json(url, opener, method="POST", filename=filename, payload=payload)
    media_id = body.get("media_id")
    if not media_id:
        raise WeComError("企业微信没有返回文件 media_id。")
    return str(media_id)


def _upload_webhook(webhook: str, filename: str, payload: bytes, opener) -> str:
    parts = urlsplit(webhook)
    query = parts.query
    if query:
        query = f"{query}&type=file"
    else:
        query = "type=file"
    upload = urlunsplit(
        (parts.scheme, parts.netloc, "/cgi-bin/webhook/upload_media", query, "")
    )
    body = _read_json(upload, opener, method="POST", filename=filename, payload=payload)
    media_id = body.get("media_id")
    if not media_id:
        raise WeComError("企业微信没有返回文件 media_id。")
    return str(media_id)


def _post_json(url: str, body: dict, opener) -> dict:
    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    request = Request(
        url,
        data=data,
        method="POST",
        headers={"Content-Type": "application/json; charset=utf-8"},
    )
    return _execute(request, opener)


def _read_json(url: str, opener, method: str, filename: str | None = None, payload: bytes | None = None) -> dict:
    if filename is None:
        request = Request(url, method=method)
    else:
        content, content_type = _multipart(filename, payload or b"")
        request = Request(
            url,
            data=content,
            method="POST",
            headers={"Content-Type": content_type},
        )
    return _execute(request, opener)


def _execute(request: Request, opener) -> dict:
    try:
        with opener(request, timeout=30) as response:
            raw = response.read()
    except HTTPError as exc:
        raw = exc.read()
        parsed = _parse_body(raw, request.full_url)
        if parsed is None:
            raise WeComError(f"企业微信请求失败（HTTP {exc.code}）：{_safe_url(request.full_url)}") from exc
        _raise_if_failed(parsed)
        return parsed
    except URLError as exc:
        raise WeComError(f"连接企业微信失败：{exc.reason}") from exc
    parsed = _parse_body(raw, request.full_url)
    if parsed is None:
        raise WeComError(f"企业微信返回的不是 JSON：{_safe_url(request.full_url)}")
    _raise_if_failed(parsed)
    return parsed


def _parse_body(raw: bytes, url: str) -> dict | None:
    try:
        parsed = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        return None
    if not isinstance(parsed, dict):
        raise WeComError(f"企业微信返回了无法识别的内容：{_safe_url(url)}")
    return parsed


def _raise_if_failed(payload: dict) -> None:
    code = payload.get("errcode", 0)
    if code in (None, 0):
        return
    message = str(payload.get("errmsg") or "未知错误")
    hint = _HINTS.get(int(code))
    text = f"企业微信返回错误 {code}：{message}。"
    if hint:
        text += hint
    raise WeComError(text)


def _multipart(filename: str, payload: bytes) -> tuple[bytes, str]:
    boundary = f"----xlsxwecom{uuid.uuid4().hex}"
    ascii_name = filename if filename.isascii() else "result.xlsx"
    quoted = _percent(filename)
    header = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="media"; filename="{ascii_name}"; '
        f"filename*=UTF-8''{quoted}\r\n"
        "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n"
        "\r\n"
    ).encode("utf-8")
    body = header + payload + f"\r\n--{boundary}--\r\n".encode("utf-8")
    return body, f"multipart/form-data; boundary={boundary}"


def _percent(value: str) -> str:
    return "".join(f"%{byte:02X}" if byte >= 128 or chr(byte) in " %\"'#" else chr(byte) for byte in value.encode("utf-8"))


def _clip(text: str) -> str:
    data = text.encode("utf-8")
    if len(data) <= _TEXT_LIMIT:
        return text
    note = "\n其余列见文件。"
    budget = _TEXT_LIMIT - len(note.encode("utf-8"))
    trimmed = data[:budget].decode("utf-8", errors="ignore")
    return trimmed + note


def _safe_url(url: str) -> str:
    parts = urlsplit(url)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, "", ""))
