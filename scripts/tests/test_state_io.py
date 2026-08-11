# 关联方识别与核查 — 状态文件原子写入测试

import json
import os
from pathlib import Path
import tempfile
import unittest

try:
    from scripts.state_io import atomic_write_json
except ModuleNotFoundError:
    atomic_write_json = None


class AtomicWriteJsonTests(unittest.TestCase):
    def _writer(self):
        self.assertIsNotNone(atomic_write_json, "尚未实现共用状态写入函数")
        return atomic_write_json

    def test_临时占用后第三次原子替换成功(self):
        target = Path(tempfile.mkdtemp(prefix="rpi_state_retry_")) / "state.json"
        attempts = []
        waits = []

        def replace(source, destination):
            attempts.append((source, destination))
            if len(attempts) < 3:
                raise PermissionError("文件暂时被占用")
            os.replace(source, destination)

        self._writer()(
            target,
            {"企业": "甲公司"},
            replace=replace,
            sleeper=waits.append,
        )

        raw = target.read_bytes()
        self.assertEqual(len(attempts), 3)
        self.assertEqual(waits, [0.2, 0.2])
        self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
        self.assertEqual(json.loads(raw.decode("utf-8")), {"企业": "甲公司"})

    def test_连续占用达到上限时保留旧状态(self):
        target = Path(tempfile.mkdtemp(prefix="rpi_state_preserve_")) / "state.json"
        target.write_text('{"status":"old"}', encoding="utf-8")
        attempts = []
        waits = []

        def replace(_source, _destination):
            attempts.append(1)
            raise PermissionError("文件持续被占用")

        with self.assertRaises(PermissionError):
            self._writer()(
                target,
                {"status": "new"},
                replace=replace,
                sleeper=waits.append,
            )

        self.assertEqual(len(attempts), 3)
        self.assertEqual(waits, [0.2, 0.2])
        self.assertEqual(target.read_text(encoding="utf-8"), '{"status":"old"}')


if __name__ == "__main__":
    unittest.main()
