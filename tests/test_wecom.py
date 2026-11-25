import json
import os
import unittest
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import mock
from urllib.error import HTTPError

from openpyxl import Workbook

from xlsx_wecom.config import SkillConfig
from xlsx_wecom.means import process_workbook
from xlsx_wecom.wecom import (
    AppCredentials,
    WeComError,
    WeComNotConfigured,
    WebhookCredentials,
    load_local_env,
    resolve_credentials,
    send_outcome,
)


def config() -> SkillConfig:
    return SkillConfig(
        label="均值",
        header=True,
        decimals=4,
        blank_cells="skip",
        sheets="active",
        output="copy",
    )


class FakeResponse:
    def __init__(self, payload: dict):
        self._raw = json.dumps(payload).encode("utf-8")

    def read(self):
        return self._raw

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


class SendTests(unittest.TestCase):
    def test_app_message_goes_to_the_configured_user_and_not_the_webhook(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "成绩.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["姓名", "语文"])
            sheet.append(["甲", 90])
            sheet.append(["乙", 80])
            workbook.save(source)
            workbook.close()
            outcome = process_workbook(source, config())
            calls = []

            def opener(request, timeout=30):
                body = request.data or b""
                calls.append((request.full_url, request.get_method(), body))
                url = request.full_url
                if "gettoken" in url:
                    return FakeResponse({"errcode": 0, "access_token": "TOKEN"})
                if "media/upload" in url:
                    self.assertIn(b"filename*=UTF-8''", body)
                    return FakeResponse({"errcode": 0, "media_id": "MID"})
                sent = json.loads(body.decode("utf-8"))
                self.assertEqual(sent["touser"], "me")
                self.assertEqual(sent["agentid"], 1001)
                if sent["msgtype"] == "text":
                    self.assertIn("85.0000", sent["text"]["content"])
                if sent["msgtype"] == "file":
                    self.assertEqual(sent["file"]["media_id"], "MID")
                return FakeResponse({"errcode": 0, "errmsg": "ok"})

            note = send_outcome(
                outcome,
                AppCredentials("corp", "secret", 1001, "me"),
                opener=opener,
            )
            self.assertIn("me", note)
            self.assertTrue(any("gettoken" in url for url, _method, _body in calls))
            self.assertFalse(any("webhook" in url for url, _method, _body in calls))
            self.assertNotIn("secret", calls[2][0])

    def test_webhook_is_used_only_when_the_app_is_not_configured(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "成绩.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["项目", "数量"])
            sheet.append(["甲", 1])
            workbook.save(source)
            workbook.close()
            outcome = process_workbook(source, config())
            webhook = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=abc"

            def opener(request, timeout=30):
                if "upload_media" in request.full_url:
                    self.assertIn("key=abc", request.full_url)
                    return FakeResponse({"errcode": 0, "media_id": "FILE"})
                return FakeResponse({"errcode": 0, "errmsg": "ok"})

            note = send_outcome(outcome, WebhookCredentials(webhook), opener=opener)
            self.assertIn("群", note)

    def test_api_error_includes_the_hint_without_the_secret(self):
        def opener(request, timeout=30):
            raise HTTPError(
                request.full_url,
                200,
                "denied",
                hdrs=None,
                fp=BytesIO(json.dumps({"errcode": 60020, "errmsg": "not allow ip"}).encode("utf-8")),
            )

        with TemporaryDirectory() as directory:
            outcome = _outcome_in(Path(directory))
            with self.assertRaises(WeComError) as caught:
                send_outcome(
                    outcome,
                    AppCredentials("corp", "top-secret", 1, "me"),
                    opener=opener,
                )
        self.assertIn("可信 IP", str(caught.exception))
        self.assertNotIn("top-secret", str(caught.exception))

    def test_resolve_prefers_the_user_over_the_group(self):
        env = {
            "WECOM_CORP_ID": "corp",
            "WECOM_SECRET": "secret",
            "WECOM_AGENT_ID": "7",
            "WECOM_USER_ID": "me",
            "WECOM_WEBHOOK": "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=abc",
        }
        with mock.patch.dict(os.environ, env, clear=True):
            credentials = resolve_credentials()
        self.assertIsInstance(credentials, AppCredentials)
        self.assertEqual(credentials.user_id, "me")

    def test_reject_broadcast_to_everyone(self):
        env = {
            "WECOM_CORP_ID": "corp",
            "WECOM_SECRET": "secret",
            "WECOM_AGENT_ID": "7",
            "WECOM_USER_ID": "@all",
        }
        with mock.patch.dict(os.environ, env, clear=True):
            with self.assertRaises(WeComError):
                resolve_credentials()

    def test_missing_credentials(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            with self.assertRaises(WeComNotConfigured):
                resolve_credentials()

    def test_env_file_does_not_override_existing_variables(self):
        with TemporaryDirectory() as directory:
            env_file = Path(directory) / ".env"
            env_file.write_text("WECOM_USER_ID=from-file\nWECOM_CORP_ID=from-file\n", encoding="utf-8")
            with mock.patch.dict(os.environ, {"WECOM_USER_ID": "already"}, clear=True):
                load_local_env(env_file)
                self.assertEqual(os.environ["WECOM_USER_ID"], "already")
                self.assertEqual(os.environ["WECOM_CORP_ID"], "from-file")


def _outcome_in(directory: Path):
    source = directory / "成绩.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["项目", "数量"])
    sheet.append(["甲", 1])
    workbook.save(source)
    workbook.close()
    return process_workbook(source, config())


if __name__ == "__main__":
    unittest.main()
