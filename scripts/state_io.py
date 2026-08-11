# 关联方识别与核查 — 共用状态文件写入

import json
import os
from pathlib import Path
import time
from typing import Callable


def atomic_write_json(
    path: Path,
    payload: dict,
    *,
    replace_attempts: int = 3,
    replace_delay_seconds: float = 0.2,
    replace: Callable[[Path, Path], None] = os.replace,
    sleeper: Callable[[float], None] = time.sleep,
) -> None:
    """以无 BOM UTF-8 写入 JSON，并有限重试 Windows 原子替换。"""
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(destination.suffix + ".tmp")
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    with open(temporary, "w", encoding="utf-8", newline="\n") as stream:
        stream.write(text)
        stream.write("\n")
        stream.flush()
        os.fsync(stream.fileno())

    for attempt in range(replace_attempts):
        try:
            replace(temporary, destination)
            return
        except PermissionError:
            if attempt + 1 >= replace_attempts:
                raise
            sleeper(replace_delay_seconds)
