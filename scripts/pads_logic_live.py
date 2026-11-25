"""Read the schematic that is already open in PADS Logic.

Attaches with GetActiveObject only. Dispatch() on the program id would start a
blank Logic, so this module never does that.
"""
from __future__ import annotations

import time
from pathlib import Path

from bomscan import REF_RE

LOGIC_NOT_RUNNING = (
    "读取失败。未检测到正在运行的 PADS Logic。\n"
    "请先打开原理图，再生成 ECO。"
)
LOGIC_NO_DOCUMENT = (
    "读取失败。PADS Logic 没有打开的原理图。\n"
    "请打开要变更的原理图后重试。"
)
LOGIC_BUSY = (
    "读取失败。PADS Logic 当前被占用，无法完成读取。\n"
    "请关闭 Logic 中的提示窗口后重试。原理图无需另行导出。"
)
LOGIC_NO_TYPE = (
    "读取失败。当前原理图没有可用的 Part Type。\n"
    "请确认活动文档是原理图，且元件带有 Part Type。"
)
LOGIC_READ_FAILED = (
    "读取失败。当前原理图未能读出。\n"
    "请确认 PADS Logic 已打开原理图，并关闭其中的提示窗口后重试。"
)
LOGIC_ACCESS_DENIED = (
    "读取失败。无法访问正在运行的 PADS Logic。\n"
    "请关闭本工具和 PADS Logic，再以相同权限重新打开："
    "两者都用普通方式，或两者都用管理员身份。"
)
LOGIC_NO_MEMBERS = (
    "读取失败。已连接到 PADS Logic，但读不到元件列表。\n"
    "请确认当前窗口打开的是原理图。"
)

_PROGIDS = ("PowerLogic.Application", "PowerLogic.Application.1")
_BUSY_HR = {-2147418111, -2147417846, -2146777998}
_ABSENT_HR = {-2147221021, -2147221164, -2147023174}
# E_ACCESSDENIED, CO_E_SERVER_EXEC_FAILURE. Win32 code 5 is handled with the text.
_ACCESS_HR = {-2147024891, -2146959355}
# DISP_E_MEMBERNOTFOUND, DISP_E_UNKNOWNNAME.
_MEMBER_HR = {-2147352573, -2147352570}
_VALUE_KEYS = {"value", "val", "值"}


class SchematicReadError(Exception):
    """User-facing read failure. The message is the whole log line."""


def _hresult(exc: BaseException) -> int | None:
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return int(args[0])
    return None


def _chain(exc: BaseException):
    seen: set[int] = set()
    cur: BaseException | None = exc
    while cur is not None and id(cur) not in seen:
        seen.add(id(cur))
        yield cur
        cur = cur.__cause__ or cur.__context__


def _is_busy(exc: BaseException) -> bool:
    for cur in _chain(exc):
        if _hresult(cur) in _BUSY_HR:
            return True
        text = str(cur)
        if "被呼叫方拒绝接收呼叫" in text or "应用程序正忙" in text:
            return True
    return False


def _is_absent(exc: BaseException) -> bool:
    return any(_hresult(cur) in _ABSENT_HR for cur in _chain(exc))


def _is_access(exc: BaseException) -> bool:
    for cur in _chain(exc):
        code = _hresult(cur)
        if code in _ACCESS_HR:
            return True
        text = str(cur)
        if "拒绝访问" in text or "Access is denied" in text:
            return True
        if code == 5 and ("GetActiveObject" in text or "OLE" in text):
            return True
    return False


def _is_member_missing(exc: BaseException) -> bool:
    for cur in _chain(exc):
        if _hresult(cur) in _MEMBER_HR:
            return True
        if "找不到成员" in str(cur):
            return True
    return False


def _read_failed(stage: str) -> str:
    if not stage:
        return LOGIC_READ_FAILED
    return (
        "读取失败。当前原理图未能读出。\n"
        f"失败环节：{stage}\n"
        "请确认 PADS Logic 已打开原理图，并关闭其中的提示窗口后重试。"
    )


def _as_logic_error(exc: BaseException, stage: str = "") -> SchematicReadError:
    if _is_busy(exc):
        return SchematicReadError(LOGIC_BUSY)
    if _is_access(exc):
        return SchematicReadError(LOGIC_ACCESS_DENIED)
    if _is_absent(exc):
        return SchematicReadError(LOGIC_NOT_RUNNING)
    if stage == "读取元件列表" and _is_member_missing(exc):
        return SchematicReadError(LOGIC_NO_MEMBERS)
    return SchematicReadError(_read_failed(stage))


def _retry_busy(fn):
    for attempt in range(4):
        try:
            return fn()
        except SchematicReadError as e:
            if attempt == 3 or str(e) != LOGIC_BUSY:
                raise
            time.sleep(0.4)


def _text(value) -> str:
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(value).strip()
    return ""


def _prop_text(obj, name: str) -> str:
    try:
        return _text(getattr(obj, name))
    except Exception:
        return ""


def _doc_name(doc) -> str:
    for prop in ("Name", "FullName"):
        text = _prop_text(doc, prop)
        if text:
            return Path(text.replace("/", "\\")).name
    return "当前原理图"


def _norm_path(path: str) -> str:
    text = (path or "").strip().strip('"')
    if not text:
        return ""
    try:
        candidate = Path(text)
        if candidate.exists():
            return str(candidate.resolve()).casefold()
    except OSError:
        pass
    return str(Path(text)).replace("/", "\\").rstrip("\\").casefold()


def _bare_name(doc) -> str:
    name = _prop_text(doc, "Name")
    if not name:
        return ""
    return Path(name.replace("/", "\\")).name


def _doc_path(doc) -> str:
    """Full path when Logic has one. A bare name is not a path."""
    full = _prop_text(doc, "FullName")
    if full and (full.startswith("\\\\") or (len(full) > 2 and full[1] == ":")):
        return full
    folder = _prop_text(doc, "Path")
    name = _prop_text(doc, "Name") or _doc_name(doc)
    if folder and name:
        return str(Path(folder) / name)
    return ""


def _same_path(left: str, right: str) -> bool:
    a, b = _norm_path(left), _norm_path(right)
    return bool(a and b and a == b)


def _refdes(comp) -> str:
    for prop in ("Name", "RefDes"):
        text = _prop_text(comp, prop)
        if text:
            return text.upper()
    return ""


def _part_type(comp) -> str:
    for prop in ("PartType", "PartTypeName"):
        try:
            raw = getattr(comp, prop)
        except Exception:
            continue
        text = _text(raw)
        if text:
            return text.split(",")[0].strip()
        if raw is None or isinstance(raw, (str, int, float, bool)):
            continue
        named = _prop_text(raw, "Name")
        if named:
            return named.split(",")[0].strip()
    return ""


def _attr_text(attr) -> str:
    if attr is None:
        return ""
    direct = _text(attr)
    if direct:
        return direct
    for field in ("Value", "Text"):
        text = _prop_text(attr, field)
        if text:
            return text
    return ""


def _attr_value(obj) -> str:
    if obj is None or isinstance(obj, (str, int, float, bool)):
        return ""
    for coll_name in ("Attributes", "Attrs"):
        try:
            coll = getattr(obj, coll_name)
        except Exception:
            continue
        if coll is None:
            continue
        item = getattr(coll, "Item", None)
        if callable(item):
            for key in ("Value", "Val", "值"):
                try:
                    text = _attr_text(item(key))
                except Exception:
                    continue
                if text:
                    return text
        try:
            members = list(coll)
        except Exception:
            members = []
        for attr in members:
            name = _prop_text(attr, "Name").lower()
            if name in _VALUE_KEYS:
                text = _attr_text(attr)
                if text:
                    return text
    return ""


def _value(comp) -> str:
    found = _attr_value(comp)
    if found:
        return found
    try:
        part = getattr(comp, "PartType")
    except Exception:
        part = None
    if part is not None and not isinstance(part, (str, int, float, bool)):
        return _attr_value(part)
    return ""


def _take_items(coll):
    """1-based Item, then 0-based. A busy collection stays a busy error."""
    count = int(coll.Count)
    try:
        return [coll.Item(i) for i in range(1, count + 1)]
    except Exception as e:
        if _is_busy(e):
            raise SchematicReadError(LOGIC_BUSY) from e
        try:
            return [coll.Item(i) for i in range(count)]
        except Exception as err:
            if _is_busy(err):
                raise SchematicReadError(LOGIC_BUSY) from err
            raise e


def _components(doc):
    try:
        comps = doc.Components
    except Exception as e:
        raise _as_logic_error(e, "读取元件列表") from e
    if comps is None:
        raise SchematicReadError(LOGIC_NO_DOCUMENT)
    try:
        return _take_items(comps)
    except SchematicReadError:
        raise
    except Exception as first_error:
        try:
            return list(comps)
        except Exception as e:
            raise _as_logic_error(first_error or e, "读取元件列表") from e


def maps_from_document(doc) -> tuple[dict[str, str], dict[str, str], str]:
    """RefDes → Part Type, RefDes → Value, and the document name."""
    bom_type: dict[str, str] = {}
    bom_val: dict[str, str] = {}
    for comp in _components(doc):
        ref = _refdes(comp)
        if not REF_RE.match(ref):
            continue
        ptype = _part_type(comp)
        if not ptype:
            continue
        bom_type[ref] = ptype
        val = _value(comp)
        if val:
            bom_val[ref] = val
    if not bom_type:
        raise SchematicReadError(LOGIC_NO_TYPE)
    return bom_type, bom_val, _doc_name(doc)


def _active_document(app):
    try:
        doc = app.ActiveDocument
    except Exception as e:
        raise _as_logic_error(e, "读取活动文档") from e
    if doc is None:
        raise SchematicReadError(LOGIC_NO_DOCUMENT)
    return doc


def _iter_collection(coll):
    if coll is None:
        return []
    try:
        return _take_items(coll)
    except SchematicReadError:
        raise
    except Exception:
        try:
            return list(coll)
        except Exception as err:
            if _is_busy(err):
                raise SchematicReadError(LOGIC_BUSY) from err
            return []


def _open_documents(app):
    try:
        coll = getattr(app, "Documents")
    except Exception as e:
        if _is_busy(e):
            raise SchematicReadError(LOGIC_BUSY) from e
        coll = None
    docs = _iter_collection(coll)
    try:
        active = app.ActiveDocument
    except Exception as e:
        if _is_busy(e):
            raise SchematicReadError(LOGIC_BUSY) from e
        active = None
    if active is not None and all(doc is not active for doc in docs):
        docs.append(active)
    return docs


def _document_for(app, expected_path: str | None):
    """Open document that matches expected_path, or the active one when no path is set."""
    if not (expected_path or "").strip():
        return _active_document(app)
    docs = _open_documents(app)
    exact = [doc for doc in docs if _same_path(_doc_path(doc), expected_path)]
    if exact:
        return exact[0]
    want = Path(expected_path).name.casefold()
    named = [
        doc for doc in docs
        if _prop_text(doc, "Name").casefold() == want or Path(_doc_path(doc) or _prop_text(doc, "Name")).name.casefold() == want
    ]
    if len(named) == 1:
        return named[0]
    active = _active_document(app)
    got = _doc_path(active) or _doc_name(active) or "（未能取得路径）"
    if len(named) > 1:
        raise SchematicReadError(
            "读取失败。有多份同名原理图，无法按读取路径确定目标。\n"
            f"读取路径：{expected_path}\n"
            "请改为完整路径后再生成。"
        )
    raise SchematicReadError(
        "读取失败。当前打开的原理图与读取路径不一致。\n"
        f"读取路径：{expected_path}\n"
        f"正在打开：{got}\n"
        "请在 PADS Logic 中打开该原理图，或修改读取路径后再生成。"
    )


def inspect_active_schematic(app=None) -> tuple[str, str]:
    """(path or document name, notice). Never raises.

    Notice stays empty when Logic is not running, so opening the tool
    does not log a failure just because Logic is closed.
    """
    try:
        application = app if app is not None else _attach()
    except SchematicReadError as e:
        text = str(e)
        if text == LOGIC_NOT_RUNNING:
            return "", ""
        return "", text
    except Exception as e:
        err = _as_logic_error(e, "连接 PADS Logic")
        if str(err) == LOGIC_NOT_RUNNING:
            return "", ""
        return "", str(err)
    try:
        doc = application.ActiveDocument
    except Exception as e:
        err = _as_logic_error(e, "读取活动文档")
        if str(err) == LOGIC_NOT_RUNNING:
            return "", ""
        return "", str(err)
    if doc is None:
        return "", ""
    return _doc_path(doc) or _bare_name(doc), ""


def peek_active_schematic_path(app=None) -> str:
    """Path or document name of the active schematic. Empty when Logic is closed.

    One attempt, no retry. Used to fill the path field; it must not raise.
    """
    path, _note = inspect_active_schematic(app)
    return path


def _attach():
    import pythoncom
    import win32com.client.dynamic as dynamic

    try:
        pythoncom.CoInitialize()
    except Exception:
        pass
    last: BaseException | None = None
    denied: BaseException | None = None
    for progid in _PROGIDS:
        try:
            # Never Dispatch(progid): that starts a blank Logic.
            unknown = pythoncom.GetActiveObject(progid)
        except Exception as e:
            last = e
            if _is_access(e):
                denied = e
            continue
        try:
            # Late binding. An early-bound makepy wrapper goes stale across
            # Logic versions and then reports Member not found.
            return dynamic.Dispatch(unknown)
        except Exception as e:
            raise _as_logic_error(e, "连接 PADS Logic") from e
    if denied is not None:
        raise _as_logic_error(denied, "连接 PADS Logic") from denied
    if last is None:
        raise SchematicReadError(LOGIC_NOT_RUNNING)
    raise _as_logic_error(last, "连接 PADS Logic") from last


def read_active_schematic(app=None, expected_path: str | None = None):
    """Return (bom_type, bom_val, doc_name, full_path) from the running Logic.

    expected_path limits the read to that schematic. Pass app only in tests.
    The product path attaches to Logic that is already open.
    """

    def once():
        try:
            application = app if app is not None else _attach()
            doc = _document_for(application, expected_path)
            bom_type, bom_val, name = maps_from_document(doc)
            return bom_type, bom_val, name, _doc_path(doc) or name
        except SchematicReadError:
            raise
        except Exception as e:
            raise _as_logic_error(e, "读取原理图") from e

    return _retry_busy(once)
