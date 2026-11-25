"""Live schematic read: fake Logic objects, no PADS install required."""
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "tool"))

import pads_logic_live
from pads_logic_live import (
    LOGIC_BUSY,
    LOGIC_NOT_RUNNING,
    LOGIC_NO_DOCUMENT,
    SchematicReadError,
    maps_from_document,
    peek_active_schematic_path,
    read_active_schematic,
)
import pads_eco_app


class _Com(Exception):
    pass


class _Attr:
    def __init__(self, name, value):
        self.Name = name
        self.Value = value


class _Bag:
    def __init__(self, attrs):
        self._attrs = list(attrs)

    def __iter__(self):
        return iter(self._attrs)


class _Part:
    def __init__(self, name):
        self.Name = name


class _Comp:
    def __init__(self, name, part, value=None):
        self.Name = name
        self.PartType = part
        attrs = [] if value is None else [_Attr("Value", value)]
        self.Attributes = _Bag(attrs)


class _Col:
    def __init__(self, comps):
        self._comps = list(comps)
        self.Count = len(self._comps)

    def Item(self, index):
        return self._comps[index - 1]


class _Docs:
    def __init__(self, docs):
        self._docs = list(docs)
        self.Count = len(self._docs)

    def Item(self, index):
        return self._docs[index - 1]


class _App:
    def __init__(self, docs, active=0):
        self.Documents = _Docs(docs)
        self.ActiveDocument = docs[active]


class _Doc:
    def __init__(self, comps, name="demo.sch", full=""):
        self.Name = name
        self.FullName = full
        self.Components = _Col(comps)


def test_component_maps_ref_type_and_value():
    doc = _Doc([
        _Comp("r12", "C_M_10P_0402,10pF", "10p"),
        _Comp("u1", _Part("CONN_USB"), "USB"),
        _Comp("GND", "C_M_10P_0402", "skip"),
    ])
    bom_type, bom_val, name = maps_from_document(doc)
    assert name == "demo.sch"
    assert bom_type["R12"] == "C_M_10P_0402"
    assert bom_val["R12"] == "10p"
    assert bom_type["U1"] == "CONN_USB"
    assert "GND" not in bom_type


def test_busy_com_is_official_sentence(monkeypatch):
    monkeypatch.setattr(time, "sleep", lambda _s: None)
    calls = {"n": 0}

    class App:
        @property
        def ActiveDocument(self):
            calls["n"] += 1
            raise _Com(-2147418111, "被呼叫方拒绝接收呼叫。", None, None)

    try:
        read_active_schematic(app=App())
    except SchematicReadError as e:
        text = str(e)
    else:
        raise AssertionError("expected SchematicReadError")
    assert text == LOGIC_BUSY
    assert calls["n"] == 4
    assert "com_error" not in text
    assert "-2147418111" not in text
    assert "被呼叫方" not in text


def test_not_running_and_no_document():
    absent = pads_logic_live._as_logic_error(_Com(-2147221021, "MK_E_UNAVAILABLE", None, None))
    assert str(absent) == LOGIC_NOT_RUNNING
    assert "-2147221021" not in str(absent)

    class App:
        ActiveDocument = None

    try:
        read_active_schematic(app=App())
    except SchematicReadError as e:
        text = str(e)
    else:
        raise AssertionError("expected SchematicReadError")
    assert text == LOGIC_NO_DOCUMENT
    assert "com_error" not in text


def test_expected_path_selects_that_schematic():
    other = _Doc([], name="a.sch", full=r"D:\boards\a.sch")
    target = _Doc([_Comp("r12", "C_M_10P_0402", "10p")], name="b.sch", full=r"D:\boards\b.sch")
    bom_type, bom_val, name, path = read_active_schematic(
        app=_App([other, target], active=0),
        expected_path=r"D:\boards\b.sch",
    )
    assert name == "b.sch"
    assert path == r"D:\boards\b.sch"
    assert bom_type["R12"] == "C_M_10P_0402"
    assert bom_val["R12"] == "10p"
    assert peek_active_schematic_path(app=_App([other, target])) == r"D:\boards\a.sch"


def test_access_denied_is_official_sentence():
    denied = pads_logic_live._as_logic_error(
        _Com(-2147024891, "OLE error 0x80070005", None, None)
    )
    text = str(denied)
    assert "无法访问正在运行的 PADS Logic" in text
    assert "相同权限" in text
    assert "com_error" not in text
    assert "2147024891" not in text
    assert "80070005" not in text

    win32 = pads_logic_live._as_logic_error(_Com(5, "GetActiveObject", "拒绝访问。"))
    assert "无法访问正在运行的 PADS Logic" in str(win32)

    launch = pads_logic_live._as_logic_error(
        _Com(-2146959355, "CO_E_SERVER_EXEC_FAILURE", None, None)
    )
    assert "无法访问正在运行的 PADS Logic" in str(launch)
    assert "2146959355" not in str(launch)


def test_active_document_access_denied_is_not_generic():
    class App:
        @property
        def ActiveDocument(self):
            raise _Com(-2147024891, "OLE error 0x80070005", None, None)

    try:
        read_active_schematic(app=App())
    except SchematicReadError as e:
        text = str(e)
    else:
        raise AssertionError("expected SchematicReadError")
    assert "无法访问正在运行的 PADS Logic" in text
    assert "当前原理图未能读出" not in text
    assert "com_error" not in text
    assert "2147024891" not in text


def test_unknown_failure_names_the_step():
    class App:
        @property
        def ActiveDocument(self):
            raise _Com(1, "weird-internal", None, None)

    try:
        read_active_schematic(app=App())
    except SchematicReadError as e:
        text = str(e)
    else:
        raise AssertionError("expected SchematicReadError")
    assert "失败环节：读取活动文档" in text
    assert "weird-internal" not in text
    assert "com_error" not in text


def test_missing_component_list_is_official_sentence():
    class Doc:
        Name = "a.sch"
        FullName = r"D:\boards\a.sch"

        @property
        def Components(self):
            raise _Com(-2147352573, "找不到成员。", None, None)

    try:
        read_active_schematic(app=_App([Doc()]))
    except SchematicReadError as e:
        text = str(e)
    else:
        raise AssertionError("expected SchematicReadError")
    assert "读不到元件列表" in text
    assert "2147352573" not in text
    assert "找不到成员" not in text
    assert "com_error" not in text


def test_peek_uses_document_name_when_unsaved():
    doc = _Doc(
        [_Comp("r12", "C_M_10P_0402", "10p")],
        name="FG560-GL-00.sch",
        full="",
    )
    assert peek_active_schematic_path(app=_App([doc])) == "FG560-GL-00.sch"


def test_zero_based_components_still_read():
    class _ZeroCol:
        def __init__(self):
            self._comps = [_Comp("r12", "C_M_10P_0402", "10p")]
            self.Count = 1

        def Item(self, index):
            if index == 0:
                return self._comps[0]
            raise _Com(-2147352567, "bad index", None, None)

    doc = _Doc([], name="a.sch", full=r"D:\boards\a.sch")
    doc.Components = _ZeroCol()
    bom_type, bom_val, name = maps_from_document(doc)
    assert name == "a.sch"
    assert bom_type["R12"] == "C_M_10P_0402"
    assert bom_val["R12"] == "10p"


def _install_com(monkeypatch, get_active, dispatch):
    import types

    dynamic = types.ModuleType("win32com.client.dynamic")
    dynamic.Dispatch = dispatch
    client = types.ModuleType("win32com.client")

    def early(*_args, **_kwargs):
        raise AssertionError("early-bound Dispatch")

    client.Dispatch = early
    client.dynamic = dynamic
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: None
    pythoncom.GetActiveObject = get_active
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setitem(sys.modules, "win32com.client.dynamic", dynamic)


def test_attach_uses_dynamic_dispatch(monkeypatch):
    seen = {}

    def get_active(progid):
        seen["progid"] = progid
        return "unk"

    def dispatch(unknown):
        seen["unknown"] = unknown
        return "app"

    _install_com(monkeypatch, get_active, dispatch)
    assert pads_logic_live._attach() == "app"
    assert seen == {"progid": "PowerLogic.Application", "unknown": "unk"}


def test_attach_reports_denial(monkeypatch):
    def get_active(_progid):
        raise _Com(-2147024891, "OLE error 0x80070005", None, None)

    _install_com(monkeypatch, get_active, lambda unknown: unknown)
    try:
        pads_logic_live._attach()
    except SchematicReadError as e:
        text = str(e)
    else:
        raise AssertionError("expected SchematicReadError")
    assert "无法访问正在运行的 PADS Logic" in text
    assert "2147024891" not in text
    assert "com_error" not in text


def test_inspect_is_quiet_when_logic_is_absent(monkeypatch):
    def get_active(_progid):
        raise _Com(-2147221021, "MK_E_UNAVAILABLE", None, None)

    _install_com(monkeypatch, get_active, lambda unknown: unknown)
    path, note = pads_logic_live.inspect_active_schematic()
    assert path == ""
    assert note == ""


def test_inspect_reports_access_denial(monkeypatch):
    def get_active(_progid):
        raise _Com(-2147024891, "OLE error 0x80070005", None, None)

    _install_com(monkeypatch, get_active, lambda unknown: unknown)
    path, note = pads_logic_live.inspect_active_schematic()
    assert path == ""
    assert "无法访问正在运行的 PADS Logic" in note
    assert "2147024891" not in note


def test_path_mismatch_is_official_sentence():
    active = _Doc([_Comp("r12", "C_M_10P_0402", "10p")], name="a.sch", full=r"D:\boards\a.sch")
    try:
        read_active_schematic(app=_App([active]), expected_path=r"D:\boards\missing.sch")
    except SchematicReadError as e:
        text = str(e)
    else:
        raise AssertionError("expected SchematicReadError")
    assert "读取路径" in text
    assert r"D:\boards\missing.sch" in text
    assert r"D:\boards\a.sch" in text
    assert "com_error" not in text
    assert "-2147" not in text


def test_do_gen_uses_live_maps(tmp_path, monkeypatch):
    seen = {}

    def fake_run(ecn, lib, out, bom, map_encoding="utf-8-sig", bom_maps=None):
        seen["bom"] = bom
        seen["maps"] = bom_maps
        out.write_bytes(b"*PADS-ECO-V3.0-MILS*\r\n")
        return []

    def fake_checks(eco, bom, lib, bom_maps=None):
        seen["check_bom"] = bom
        seen["check_maps"] = bom_maps
        return []

    monkeypatch.setattr(pads_eco_app.ecn_to_pads_eco, "run", fake_run)
    monkeypatch.setattr(pads_eco_app.verify_eco, "run_checks", fake_checks)
    maps = ({"R12": "C_M_10P_0402"}, {"R12": "10p"})
    logs = []
    ok = pads_eco_app.do_gen(
        tmp_path / "a.xlsm",
        None,
        tmp_path / "lib.xlsx",
        tmp_path / "a.eco",
        logs.append,
        bom_maps=maps,
        bom_label="demo.sch",
    )
    text = "\n".join(logs)
    assert ok is True
    assert seen["bom"] is None
    assert seen["maps"] == maps
    assert seen["check_bom"] is None
    assert seen["check_maps"] == maps
    assert "原理图: demo.sch · 1 个位号" in text
    assert "BOM :" not in text


def test_file_bom_still_required_without_maps():
    from bomscan import schematic_maps

    try:
        schematic_maps(None, None)
    except SystemExit as e:
        text = str(e)
    else:
        raise AssertionError("expected SystemExit")
    assert "缺少原理图 BOM" in text


def test_maps_do_not_scan_a_file(monkeypatch):
    import bomscan

    def boom(_path):
        raise AssertionError("scan_bom should not run")

    monkeypatch.setattr(bomscan, "scan_bom", boom)
    bom_type, bom_val = bomscan.schematic_maps(
        None, ({"R12": "C_M_10P_0402"}, {"R12": "10p"}),
    )
    assert bom_type["R12"] == "C_M_10P_0402"
    assert bom_val["R12"] == "10p"
