"""One-click reverse check against a copy of the department ECN workbook.

The original xlsm is never opened for writing. A sibling ``原名.反向检查.xlsm``
is the file whose EBOM is replaced and whose macros run. The changelist on that
copy is saved only when Generator finishes and leaves the RF_ECN_Generation
header in place. A macro that stops early can already have cleared that sheet
in memory; saving then would wipe the previous ECN.
"""
from __future__ import annotations

import shutil
import threading
from pathlib import Path

from bomscan import ebom_rows_from_bom

# MsgBox texts that mean this run must not be reported as an empty ECN.
_BLOCKING = (
    "请先在表Check_Package",
    "不要试图修改系统时间",
    "需要整行删除",
    "ECN生成失败",
    "位号输入格式有误",
    "匹配输入有误",
    "有重复位号",
    "请填写BID",
    "没有找到替换料",
    "封装不一致",
    "封装与原理图不一致",
    "运行时错误",
    "执行失败",
    "CO-PAD",
    "EBOM中没有找到",
)
_OK_DIALOG = "执行完毕"


def interpret(finished: bool, refs: list[str], dialogs: list[str]) -> tuple[bool, str]:
    """PASS only when Generator finished and left no change rows."""
    blob = "\n".join(dialogs)
    blocked = next((s for s in _BLOCKING if s in blob), None)
    if finished and not refs and blocked is None:
        return True, "ECN 为空"
    if blocked and not finished:
        return False, "宏中途退出，ECN 未重写"
    if finished and refs:
        return False, f"ECN 仍有 {len(refs)} 行"
    if finished and blocked:
        return False, "宏有报错，不把 ECN 判成空"
    if dialogs:
        return False, "宏中途退出，ECN 未重写"
    return False, "宏没有生成完，ECN 未重写"


class ReverseReport:
    def __init__(
        self,
        removed_at: int,
        removed_blank: int,
        removed_name: int,
        kept: int,
        ebom_saved: bool,
        finished: bool,
        refs: list[str],
        dialogs: list[str],
        check_note: str,
    ):
        self.removed_at = removed_at
        self.removed_blank = removed_blank
        self.removed_name = removed_name
        self.kept = kept
        self.ebom_saved = ebom_saved
        self.finished = finished
        self.refs = refs
        self.dialogs = dialogs
        self.check_note = check_note
        self.workbook: Path | None = None

    @property
    def passed(self) -> bool:
        return interpret(self.finished, self.refs, self.dialogs)[0]

    @property
    def summary(self) -> str:
        return interpret(self.finished, self.refs, self.dialogs)[1]


def _norm(value):
    if value is None:
        return None
    if hasattr(value, "year") and hasattr(value, "hour"):
        return (value.year, value.month, value.day, value.hour, value.minute, int(value.second))
    return str(value).strip()


def _as_text(value):
    """Excel otherwise turns \"0402\" into 402 when the cell format is General."""
    if value is None:
        return None
    if isinstance(value, str):
        return value if value.strip() else None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    return str(value)


def _rectangle(rows: list[list]) -> tuple:
    width = max(len(r) for r in rows)
    out = []
    for r_i, row in enumerate(rows):
        line = []
        for c in range(width):
            value = row[c] if c < len(row) else None
            if r_i > 0:
                value = _as_text(value)
            elif isinstance(value, str) and value.strip() == "":
                value = None
            line.append(value)
        out.append(tuple(line))
    return tuple(out)


class _OkClicker:
    """Click OK on this Excel process's message boxes. Never clicks Yes/No."""

    def __init__(self, pid: int):
        self.pid = pid
        self.dialogs: list[str] = []
        self._stop = threading.Event()
        self._seen: set[str] = set()
        self._thread = threading.Thread(target=self._loop, name="ecn-msgbox", daemon=True)

    def start(self) -> None:
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        self._thread.join(timeout=2)

    def _loop(self) -> None:
        import win32con
        import win32gui

        from xlsx_open import _pid_from_hwnd

        while not self._stop.wait(0.05):
            boxes: list[int] = []

            def top(hwnd, _):
                try:
                    if (
                        win32gui.IsWindowVisible(hwnd)
                        and win32gui.GetClassName(hwnd) == "#32770"
                        and _pid_from_hwnd(hwnd) == self.pid
                    ):
                        boxes.append(hwnd)
                except Exception:
                    pass
                return True

            try:
                win32gui.EnumWindows(top, None)
            except Exception:
                continue
            for hwnd in boxes:
                self._click(win32gui, win32con, hwnd)

    def _click(self, win32gui, win32con, hwnd) -> None:
        statics: list[str] = []
        buttons: list[tuple[str, int]] = []

        def child(ch, _):
            try:
                kind = win32gui.GetClassName(ch)
                text = win32gui.GetWindowText(ch).strip()
                if kind == "Static" and text:
                    statics.append(text)
                elif kind == "Button" and text:
                    buttons.append((text, ch))
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(hwnd, child, None)
        except Exception:
            return
        body = max(statics, key=len) if statics else ""
        target = None
        for label, handle in buttons:
            if label in ("确定", "OK", "Ok"):
                target = handle
                break
        if target is None:
            ok = win32gui.GetDlgItem(hwnd, 1)  # IDOK
            if ok:
                target = ok
        if target is None and "运行时错误" in body:
            for label, handle in buttons:
                if label == "结束":
                    target = handle
                    break
        if target is None:
            return
        if body and body not in self._seen:
            self._seen.add(body)
            self.dialogs.append(body)
        try:
            win32gui.SendMessage(target, win32con.BM_CLICK, 0, 0)
        except Exception:
            pass


def _sheet(wb, name: str):
    for i in range(1, int(wb.Worksheets.Count) + 1):
        sh = wb.Worksheets(i)
        if str(sh.Name).lower() == name.lower():
            return sh
    raise RuntimeError(f"这份工作簿没有「{name}」表，不是 ECN 工具")


def _header(gen) -> tuple:
    return tuple(_norm(gen.Cells(1, col).Value) for col in (4, 7, 8, 11)) + (
        _norm(gen.Cells(2, 1).Value),
    )


def _refs(gen) -> list[str]:
    xl_up = -4162
    last = int(gen.Cells(gen.Rows.Count, 1).End(xl_up).Row)
    if last < 4:
        return []
    block = gen.Range(gen.Cells(4, 1), gen.Cells(last, 1)).Value
    if block is None:
        return []
    if not isinstance(block, tuple):
        block = ((block,),)
    refs: list[str] = []
    for row in block:
        if not isinstance(row, tuple):
            row = (row,)
        value = row[0]
        if value is not None and str(value).strip():
            refs.append(str(value).strip())
    return refs


def _write_ebom(ebom, rows: list[list]) -> None:
    data = _rectangle(rows)
    n_rows = len(data)
    n_cols = len(data[0])
    ebom.UsedRange.UnMerge()
    ebom.UsedRange.Clear()
    dest = ebom.Range(ebom.Cells(1, 1), ebom.Cells(n_rows, n_cols))
    # Text format first. A General cell stores the digits 0402 as the number 402.
    dest.NumberFormat = "@"
    dest.Value = data


def ecn_working_copy(src: Path | str) -> Path:
    """Copy the ECN tool to ``原名.反向检查.xlsm``. The source file is not opened.

    A file that is already that backup is returned as-is, so a second pass does
    not nest another suffix. Each pass that starts from the original overwrites
    the previous backup with a fresh copy.
    """
    src_path = Path(src).resolve()
    if src_path.suffix.lower() != ".xlsm":
        raise RuntimeError(f"反向检查要带宏的 ECN 工具 (.xlsm)，这是 {src_path.name}")
    if not src_path.exists():
        raise RuntimeError(f"找不到 ECN 工具: {src_path.name}")
    if src_path.stem.endswith(".反向检查"):
        return src_path
    dest = src_path.with_name(f"{src_path.stem}.反向检查.xlsm")
    if dest.resolve() == src_path:
        raise RuntimeError("备份路径与原文件相同")
    try:
        shutil.copy2(src_path, dest)
    except OSError as e:
        raise RuntimeError(f"复制 ECN 工具失败，原文件未改: {e}") from e
    return dest


def run_reverse(bom: Path | str, xlsm: Path | str, log=None) -> ReverseReport:
    rows, removed_at, removed_blank, removed_name = ebom_rows_from_bom(bom)
    kept = max(0, len(rows) - 1)

    def say(text: str) -> None:
        if log:
            log(text)

    src = Path(xlsm)
    backup = ecn_working_copy(src)
    if backup.resolve() == src.resolve():
        say(f"使用已有备份 {backup.name}")
    else:
        say(f"已复制 {backup.name}，原文件 {src.name} 不改")
    say(f"预处理  删除 值带@ {removed_at} 行，值空白 {removed_blank} 行，名称空白 {removed_name} 行，保留 {kept} 行")
    report = _drive(backup, rows, removed_at, removed_blank, removed_name, kept, say)
    report.workbook = backup
    return report


def _drive(xlsm: Path, rows, removed_at, removed_blank, removed_name, kept, say) -> ReverseReport:
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise RuntimeError("反向检查要调用 Excel 宏，本机没有 pywin32") from e

    from xlsx_open import _pid_from_hwnd, _reap

    pythoncom.CoInitialize()
    excel = None
    wb = None
    pid = 0
    clicker = None
    saved = False
    finished = False
    refs: list[str] = []
    check_note = ""
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        try:
            excel.EnableEvents = False
        except Exception:
            pass
        # 1 = msoAutomationSecurityLow. Must be set before Open, or the macros
        # we are about to run are disabled. 3 (ForceDisable) is for reading only.
        excel.AutomationSecurity = 1
        try:
            pid = _pid_from_hwnd(int(excel.Hwnd))
        except Exception:
            pid = 0
        if not pid:
            raise RuntimeError("找不到 Excel 进程，无法关掉宏的提示框")
        clicker = _OkClicker(pid)
        clicker.start()
        try:
            wb = excel.Workbooks.Open(
                str(xlsm.resolve()),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
            )
        except Exception as e:
            from xlsx_open import user_file_error

            plain = user_file_error(e)
            if plain:
                raise RuntimeError(plain) from e
            raise RuntimeError(f"打不开 ECN 工具（可能已在 Excel 中打开）: {e}") from e
        if bool(wb.ReadOnly):
            raise RuntimeError("ECN 工具已在 Excel 中打开，或文件只读。请先关掉该工作簿。")

        ebom = _sheet(wb, "EBOM")
        gen = _sheet(wb, "RF_ECN_Generation")
        check = _sheet(wb, "Check_Package")
        say("正在替换 EBOM…")
        _write_ebom(ebom, rows)
        wb.Save()
        saved = True
        say(f"EBOM 已写入备份  {xlsm.name}")

        say("正在做位号检查…")
        wb.Activate()
        excel.Run("SearchData_Global")
        check_note = str(check.Cells(2, 9).Value or "").strip()
        before = _header(gen)
        say("正在生成 ECN…")
        excel.Run("Generator")
        after = _header(gen)
        # formatChangeList rewrites row 1 and puts Ref.Des in A2. ErrorWarning
        # clears the whole sheet afterwards on a bad ref, so A2 must still be there.
        finished = after != before and after[4] == "Ref.Des" and after[0] is not None
        if finished:
            refs = _refs(gen)
            wb.Save()
            say("ECN 已生成并保存")
        else:
            say("宏没有生成完，ECN 表保持原样（未保存中途清空）")
        clicker.stop()
        dialogs = list(clicker.dialogs)
        return ReverseReport(
            removed_at,
            removed_blank,
            removed_name,
            kept,
            saved,
            finished,
            refs,
            dialogs,
            check_note,
        )
    except Exception as e:
        if saved:
            raise RuntimeError(f"备份 {xlsm.name} 里 EBOM 已写入，原文件未改。生成没有完成: {e}") from e
        raise
    finally:
        if clicker is not None:
            clicker.stop()
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        wb = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        excel = None
        import gc

        gc.collect()
        pythoncom.CoUninitialize()
        if pid:
            _reap(pid)
