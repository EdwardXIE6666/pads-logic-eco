"""Open .xlsx/.xlsm for reading, including 亿赛通 (E-SafeNet) encrypted files.

E-SafeNet decrypts per *process image name*. python.exe is usually trusted; a
packaged pads-eco.exe is not, so its bytes read back as ciphertext. For those
files we ask Excel (a trusted process) to open the workbook and hand the cell
values over COM — nothing is written to disk, so nothing gets re-encrypted.

    from xlsx_open import load_workbook_any
    wb = load_workbook_any(path)          # values only (data_only)
"""
from __future__ import annotations

import io
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


class EncryptedXlsxError(RuntimeError):
    pass


# Excel 忙时拒绝 COM。界面只说人话，不把这段代码给用户看。
# RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER / VBA_E_IGNORE
_BUSY_HR = {-2147418111, -2147417846, -2146777998}
EXCEL_BUSY_MESSAGE = (
    "读取失败。Excel 当前被占用，无法完成读取。\n"
    "请关闭 Excel 中的提示窗口，或退出所有 Excel 进程后重试。输入文件无需修改。"
)


def _hresult(exc) -> int | None:
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return int(args[0])
    return None


def _error_is_excel_busy(exc: BaseException) -> bool:
    seen: set[int] = set()
    cur: BaseException | None = exc
    while cur is not None and id(cur) not in seen:
        seen.add(id(cur))
        if _hresult(cur) in _BUSY_HR:
            return True
        text = str(cur)
        if "被呼叫方拒绝接收呼叫" in text or "应用程序正忙" in text:
            return True
        cur = cur.__cause__ or cur.__context__
    return False


def user_file_error(exc: BaseException) -> str | None:
    """Excel 忙时返回给用户看的说明。其他错误返回 None。"""
    if _error_is_excel_busy(exc):
        return EXCEL_BUSY_MESSAGE
    return None


def _as_read_error(exc: Exception) -> EncryptedXlsxError:
    if _error_is_excel_busy(exc):
        return EncryptedXlsxError(EXCEL_BUSY_MESSAGE)
    return EncryptedXlsxError("文件有加密，Excel 没有读出来。请用 Excel 打开这份文件确认正常后，再试一次。")


def _retry_busy(fn):
    """Excel 刚启动或被加载项占着时，过一会儿再叫一次。"""
    import time

    for attempt in range(4):
        try:
            return fn()
        except EncryptedXlsxError as e:
            if attempt == 3 or not _error_is_excel_busy(e):
                raise
            time.sleep(0.4)


def _quiet_excel_close(wb, excel) -> None:
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        if excel is not None:
            excel.Quit()
    except Exception:
        pass


def encryption_mode(path: Path) -> str:
    """'PLAIN' (zip header), 'LOCK' (E-SafeNet container), or 'UNKNOWN'."""
    with open(path, "rb") as f:
        head = f.read(512)
    if head[:2] == b"PK":
        return "PLAIN"
    if b"E-SafeNet" in head or b"ESafeNet" in head:
        return "LOCK"
    return "UNKNOWN"


def _via_excel_com(path: Path) -> Workbook:
    return _retry_busy(lambda: _via_excel_com_once(path))


def _via_excel_com_once(path: Path) -> Workbook:
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:  # pragma: no cover
        raise EncryptedXlsxError("文件为亿赛通加密，且缺少 pywin32，无法调用 Excel 解密读取") from e

    pythoncom.CoInitialize()
    excel = None
    wb_x = None
    pid = 0
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            pid = _pid_from_hwnd(int(excel.Hwnd))
        except Exception:
            pid = 0
        try:
            excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable: no macros
        except Exception:
            pass
        wb_x = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        out = Workbook()
        out.remove(out.active)
        n_sheets = int(wb_x.Worksheets.Count)
        for k in range(1, n_sheets + 1):
            sh = wb_x.Worksheets(k)
            ws = out.create_sheet(title=str(sh.Name)[:31])
            ur = sh.UsedRange
            vals = ur.Value
            r0, c0 = int(ur.Row), int(ur.Column)
            del ur, sh  # release COM refs promptly so Excel can exit
            if vals is None:
                continue
            if not isinstance(vals, tuple):  # single cell
                vals = ((vals,),)
            for i, row in enumerate(vals):
                if not isinstance(row, tuple):
                    row = (row,)
                for j, v in enumerate(row):
                    if v is None:
                        continue
                    if isinstance(v, str):
                        v = ILLEGAL_CHARACTERS_RE.sub("", v)
                    elif hasattr(v, "year") and hasattr(v, "hour"):  # pywintypes datetime
                        import datetime as _dt

                        v = _dt.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second)
                    ws.cell(row=r0 + i, column=c0 + j, value=v)
        buf = io.BytesIO()
        out.save(buf)
        buf.seek(0)
        return load_workbook(buf, data_only=True)
    except EncryptedXlsxError:
        raise
    except Exception as e:
        raise _as_read_error(e) from e
    finally:
        import gc

        _quiet_excel_close(wb_x, excel)
        wb_x = None
        excel = None
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        if pid:
            _reap(pid)


def _pid_from_hwnd(hwnd: int) -> int:
    import ctypes

    pid = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return int(pid.value)


def _reap(pid: int, wait_s: float = 2.0) -> None:
    """Our own hidden Excel instance sometimes survives Quit() (DLP add-ins hold it). Terminate it."""
    import ctypes
    import time

    k32 = ctypes.windll.kernel32
    SYNCHRONIZE, PROCESS_TERMINATE, PROCESS_QUERY_LIMITED = 0x00100000, 0x0001, 0x1000
    h = k32.OpenProcess(SYNCHRONIZE | PROCESS_TERMINATE | PROCESS_QUERY_LIMITED, False, pid)
    if not h:
        return
    try:
        if k32.WaitForSingleObject(h, int(wait_s * 1000)) == 0:  # exited on its own
            return
        k32.TerminateProcess(h, 0)
    finally:
        k32.CloseHandle(h)


def _used_range_rows(vals) -> list[list[str]]:
    if vals is None:
        return []
    if not isinstance(vals, tuple):
        vals = ((vals,),)
    rows: list[list[str]] = []
    for row in vals:
        if not isinstance(row, tuple):
            row = (row,)
        cells = ["" if v is None else (v if isinstance(v, str) else str(v)) for v in row]
        while cells and cells[-1] == "":
            cells.pop()
        if any(cells):
            rows.append(cells)
    return rows


def read_csv_via_excel(path: Path, accept) -> list[list[str]]:
    return _retry_busy(lambda: _read_csv_via_excel_once(path, accept))


def _read_csv_via_excel_once(path: Path, accept) -> list[list[str]]:
    """Open a CSV in Excel and return cell text.

    pads-eco.exe is not an E-SafeNet trusted process, so an encrypted CSV reads
    back as ciphertext. Excel is trusted: it decrypts, and QueryTable forces every
    column to text so Part Numbers are not rounded. Nothing is saved.
    ``accept(rows)`` picks the first parse (GBK/UTF-8, comma/tab) that looks right.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:  # pragma: no cover
        raise EncryptedXlsxError("CSV 为亿赛通加密，且缺少 pywin32，无法调用 Excel") from e

    pythoncom.CoInitialize()
    excel = None
    wb = None
    pid = 0
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            pid = _pid_from_hwnd(int(excel.Hwnd))
        except Exception:
            pid = 0
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        last: list[list[str]] = []
        src = str(Path(path).resolve())
        for origin, comma, tab in ((936, True, False), (65001, True, False), (936, False, True), (65001, False, True)):
            count = int(ws.QueryTables.Count)
            for i in range(count, 0, -1):
                ws.QueryTables(i).Delete()
            ws.Cells.Clear()
            qt = ws.QueryTables.Add(Connection="TEXT;" + src, Destination=ws.Range("A1"))
            qt.TextFilePlatform = origin
            qt.TextFileParseType = 1  # xlDelimited
            qt.TextFileStartRow = 1
            qt.TextFileCommaDelimiter = bool(comma)
            qt.TextFileTabDelimiter = bool(tab)
            qt.TextFileSemicolonDelimiter = False
            qt.TextFileSpaceDelimiter = False
            # String property. False becomes "False" and Excel splits on the letter F.
            qt.TextFileOtherDelimiter = ""
            qt.TextFileColumnDataTypes = tuple([2] * 40)  # xlTextFormat
            qt.TextFilePromptOnRefresh = False
            qt.AdjustColumnWidth = False
            qt.Refresh(False)
            last = _used_range_rows(ws.UsedRange.Value)
            qt.Delete()
            if last and accept(last):
                return last
        header = last[0] if last else []
        seen = " | ".join(str(h).strip() for h in header if str(h).strip()) or "(空)"
        raise ValueError(f"Excel 已打开这份 CSV，但没找到元件类型和 Part Number。实际表头: {seen}")
    except (EncryptedXlsxError, ValueError):
        raise
    except Exception as e:
        raise _as_read_error(e) from e
    finally:
        import gc

        _quiet_excel_close(wb, excel)
        wb = None
        excel = None
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        if pid:
            _reap(pid)


def load_workbook_any(path: str | Path):
    """openpyxl workbook (data_only=True) for a plain or E-SafeNet-encrypted xlsx/xlsm."""
    p = Path(path)
    mode = encryption_mode(p)
    if mode == "PLAIN":
        return load_workbook(p, data_only=True)
    if mode == "LOCK":
        return _via_excel_com(p)
    # Unknown header: could be transparent-encrypted for python.exe but not for us; try both.
    try:
        return load_workbook(p, data_only=True)
    except Exception:
        return _via_excel_com(p)
