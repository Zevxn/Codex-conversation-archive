import json
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import patch

import codex_bark_hook as hook


# SECTION: 禁通知时间测试

class NotificationQuietTimeTests(unittest.TestCase):
    def local_time(self, hour: int, minute: int) -> datetime:
        return datetime(2026, 9, 23, hour, minute, tzinfo=timezone(timedelta(hours=8)))

    def test_same_day_interval_includes_start_and_excludes_end(self) -> None:
        with patch.object(hook, "NOTIFICATION_QUIET_START_TIME", "00:30"), patch.object(
            hook, "NOTIFICATION_QUIET_END_TIME", "08:00"
        ):
            for hour, minute, expected in [
                (0, 29, False),
                (0, 30, True),
                (7, 59, True),
                (8, 0, False),
            ]:
                with self.subTest(hour=hour, minute=minute):
                    self.assertEqual(hook.is_notification_quiet_time(self.local_time(hour, minute)), expected)

    def test_interval_can_cross_midnight(self) -> None:
        with patch.object(hook, "NOTIFICATION_QUIET_START_TIME", "23:00"), patch.object(
            hook, "NOTIFICATION_QUIET_END_TIME", "08:00"
        ):
            for hour, minute, expected in [
                (22, 59, False),
                (23, 0, True),
                (0, 30, True),
                (7, 59, True),
                (8, 0, False),
            ]:
                with self.subTest(hour=hour, minute=minute):
                    self.assertEqual(hook.is_notification_quiet_time(self.local_time(hour, minute)), expected)

    def test_invalid_time_configuration_is_rejected(self) -> None:
        cases = [
            {"notification_quiet_start_time": "00:30"},
            {"notification_quiet_start_time": "00:30", "notification_quiet_end_time": "00:30"},
            {"notification_quiet_start_time": "24:00", "notification_quiet_end_time": "08:00"},
        ]
        with tempfile.TemporaryDirectory() as directory:
            config_path = Path(directory) / "config.json"
            for case in cases:
                with self.subTest(case=case):
                    config_path.write_text(json.dumps(case), encoding="utf-8")
                    with self.assertRaises(ValueError):
                        hook.load_configuration(config_path)

    def test_quiet_time_skips_bark_but_keeps_archive(self) -> None:
        response_time = self.local_time(0, 30)
        with tempfile.TemporaryDirectory() as directory:
            cache_path = Path(directory) / "missing-cache.json"
            archive_path = Path(directory) / "archive.json"
            record = {"record_id": "test", "duration_seconds": 120, "model": "test-model"}
            with (
                patch.object(hook, "NOTIFICATION_QUIET_START_TIME", "00:30"),
                patch.object(hook, "NOTIFICATION_QUIET_END_TIME", "08:00"),
                patch.object(hook, "ENABLE_BARK_NOTIFICATION", True),
                patch.object(hook, "ENABLE_LOCAL_HISTORY", True),
                patch.object(hook, "now_local", return_value=response_time),
                patch.object(hook, "cache_file_for_event", return_value=cache_path),
                patch.object(hook, "get_assistant_message", return_value="回答"),
                patch.object(hook, "load_prompt_cache", return_value=({}, cache_path)),
                patch.object(hook, "build_conversation_record", return_value=record),
                patch.object(hook, "append_monthly_history", return_value=archive_path) as archive,
                patch.object(hook, "send_bark") as send,
                patch.object(hook, "emit_hook_result"),
                patch.object(hook, "write_log"),
            ):
                self.assertEqual(hook.handle_stop({}), 0)
                archive.assert_called_once_with(record, response_time)
                send.assert_not_called()

    def test_manual_bark_test_also_respects_quiet_time(self) -> None:
        with (
            patch.object(hook, "NOTIFICATION_QUIET_START_TIME", "00:30"),
            patch.object(hook, "NOTIFICATION_QUIET_END_TIME", "08:00"),
            patch.object(hook, "ENABLE_BARK_NOTIFICATION", True),
            patch.object(hook, "now_local", return_value=self.local_time(7, 59)),
            patch.object(hook, "send_bark") as send,
            patch.object(hook, "emit_hook_result") as emit,
            patch.object(hook, "write_log"),
        ):
            self.assertEqual(hook.run_test_mode(), 0)
            send.assert_not_called()
            self.assertIn("未发送", emit.call_args.args[0]["systemMessage"])


# !SECTION: 禁通知时间测试

if __name__ == "__main__":
    unittest.main()
