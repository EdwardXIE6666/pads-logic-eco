"""Excel 忙时的界面说明里不能出现 COM 代码。"""
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "tool"))

import xlsx_open
from xlsx_open import EXCEL_BUSY_MESSAGE, EncryptedXlsxError, user_file_error
import pads_eco_app


class _Com(Exception):
    pass


def _busy():
    return _Com(-2147418111, "被呼叫方拒绝接收呼叫。", None, None)


def test_busy_message_has_no_code():
    text = user_file_error(_busy())
    assert text == EXCEL_BUSY_MESSAGE
    assert "com_error" not in text
    assert "-2147418111" not in text
    assert "被呼叫方" not in text
    assert "读取失败" in text
    assert "无需修改" in text


def test_wrapped_busy_message():
    try:
        raise _busy()
    except _Com as e:
        err = xlsx_open._as_read_error(e)
        err.__cause__ = e
    text = user_file_error(err)
    assert text == EXCEL_BUSY_MESSAGE
    assert "com_error" not in str(err)


def test_other_errors_stay():
    assert user_file_error(EncryptedXlsxError("缺列")) is None
    assert user_file_error(OSError("permission")) is None


def test_retry_only_while_busy(monkeypatch):
    monkeypatch.setattr(time, "sleep", lambda _s: None)
    calls = {"n": 0}

    def fn():
        calls["n"] += 1
        if calls["n"] < 3:
            err = xlsx_open._as_read_error(_busy())
            raise err from _busy()
        return "ok"

    assert xlsx_open._retry_busy(fn) == "ok"
    assert calls["n"] == 3


def test_do_gen_logs_plain_busy(tmp_path, monkeypatch):
    def boom(*_a, **_k):
        raise _busy()

    monkeypatch.setattr(pads_eco_app.ecn_to_pads_eco, "run", boom)
    logs = []
    ok = pads_eco_app.do_gen(
        tmp_path / "a.xlsm",
        tmp_path / "b.xlsx",
        tmp_path / "c.xlsx",
        tmp_path / "o.eco",
        logs.append,
    )
    text = "\n".join(logs)
    assert ok is False
    assert "读取失败" in text
    assert "com_error" not in text
    assert "-2147418111" not in text
    assert "请确认 ECN" not in text
