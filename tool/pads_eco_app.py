#!/usr/bin/env python3
"""pads-eco — PADS Logic ECO 生成工具（给工程师用，不依赖 AI）。

pads-eco.exe 双击 → 图形界面。（源码模式下带参数仍可走命令行：gen / check，见 cli()。）

规则与 SKILL.md 完全一致：scripts/ecn_to_pads_eco.py 生成，scripts/verify_eco.py 校验。
校验不通过则不留下 .eco，只打印失败原因。
生成时读取用户选择的原理图 BOM，不直接读取 PADS Logic。
反向检查另要一份导入 ECO 之后重新导出的原理图 BOM。

器件库是 exe 同目录 library\\Electronic_components_library.xlsx，没有则用打包内置的那一份。
更新器件库时选 CSV，抽出元件类型和 Part Number 写入这份 xlsx。
明文 CSV 按文本读取。若 CSV 仍是亿赛通密文，则由本机 Excel 在后台读出单元格，不另存原文件。
亿赛通(E-SafeNet)加密的 xlsx（ECN / BOM / 器件库）：自动识别，改由本机 Excel 经 COM 读取。
反向检查先把 ECN 工具复制为同目录「原名.反向检查.xlsm」，只改这份备份。
"""
from __future__ import annotations

import argparse
import contextlib
import io
import os
import re
import sys
import traceback
from pathlib import Path

# scripts/ 与本文件同包（PyInstaller 打包时通过 --paths scripts 找到）
HERE = Path(__file__).resolve().parent
SCRIPTS = HERE.parent / "scripts"
if SCRIPTS.exists() and str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import ecn_to_pads_eco  # noqa: E402
import verify_eco  # noqa: E402
from ecn_reverse import run_reverse  # noqa: E402
from xlsx_open import EncryptedXlsxError, user_file_error  # noqa: E402

LIB_NAME = "Electronic_components_library.xlsx"
VERSION = "3.5"


# ---------------------------------------------------------------------------
# 库定位
# ---------------------------------------------------------------------------
def exe_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return HERE.parent


def bundled_dir() -> Path:
    base = getattr(sys, "_MEIPASS", None)
    return Path(base) if base else HERE.parent


def default_lib() -> Path:
    for base in (exe_dir(), bundled_dir()):
        xlsx_path = base / "library" / LIB_NAME
        if xlsx_path.exists():
            return xlsx_path
    return bundled_dir() / "library" / LIB_NAME


# ---------------------------------------------------------------------------
# 核心动作（CLI 与 GUI 共用）
# ---------------------------------------------------------------------------
def do_check(eco: Path, bom: Path | None, lib: Path, log, label: str | None = None, bom_maps=None) -> bool:
    shown = label or eco.name
    try:
        errors = verify_eco.run_checks(eco, bom, lib, bom_maps=bom_maps)
    except Exception as e:
        plain = user_file_error(e)
        if plain:
            log(f"错误: {plain}")
        else:
            log(f"错误: 无法读取输入文件 — {type(e).__name__}: {e}")
        return False
    if errors:
        log(f"FAIL  {len(errors)} 项未通过  ({shown})")
        for e in errors:
            log(f"  - {e}")
        return False
    log(f"PASS  全部检查通过  ({shown})")
    return True


def _discard_generated(eco: Path, log) -> None:
    for p in (eco, eco.with_suffix(".map.csv")):
        try:
            if p.exists():
                p.unlink()
        except OSError as e:
            log(f"错误: 无法删除临时文件 {p.name} — {e}")


def _eco_run(value):
    """Accept EcoRun from the generator, or a skip list from tests that mock run()."""
    skipped = getattr(value, "skipped", None)
    if skipped is None:
        skipped = value or []
    pkg_changes = getattr(value, "pkg_changes", None) or []
    return list(skipped), list(pkg_changes)


def _user_gen_log(raw: str) -> str:
    """Keep counts and warnings. Drop library-audit internals and zero skip/pkg lines."""
    keep = []
    skipping = False
    for line in (raw or "").splitlines():
        if line.startswith("库自查:"):
            skipping = True
            continue
        if skipping:
            if not line.strip() or line[:1] in " 　" or line.startswith("…"):
                continue
            skipping = False
        m = re.match(r"modify (\d+) modify-via-NC (\d+) delete (\d+) add (\d+)", line)
        if m:
            keep.append(
                f"修改 {int(m.group(1)) + int(m.group(2))}，删除 {m.group(3)}，新增 {m.group(4)}"
            )
            continue
        k = re.match(r"skip\s+(\d+)", line)
        if k:
            if int(k.group(1)):
                keep.append(f"跳过 {k.group(1)} 个位号，原因见对照表")
            continue
        p = re.match(r"pkg-change\s+(\d+)", line)
        if p:
            if int(p.group(1)):
                keep.append(f"封装变化 {p.group(1)} 个位号")
            continue
        keep.append(line)
    return "\n".join(keep).strip()


def _log_gen_stdout(log, raw: str) -> None:
    text = _user_gen_log(raw)
    if text:
        log(text)


def do_gen(ecn: Path, bom: Path | None, lib: Path, out: Path | None, log, confirm=None, bom_maps=None, bom_label: str | None = None, confirm_pkg=None) -> bool:
    out = out or ecn.with_suffix(".eco")
    # Write beside the destination, and only replace it after verify passes.
    tmp = out.with_name(f".{out.stem}.{os.getpid()}.tmp.eco")
    log(f"ECN : {ecn}")
    if bom_maps is not None:
        log(f"原理图: {bom_label or '当前原理图'} · {len(bom_maps[0])} 个位号")
    else:
        log(f"BOM : {bom}")
    log(f"库  : {lib}")
    log(f"输出: {out}")
    log("")
    buf = io.StringIO()
    skipped = []
    pkg_changes = []
    try:
        with contextlib.redirect_stdout(buf):
            if bom_maps is not None:
                result = ecn_to_pads_eco.run(ecn, lib, tmp, bom, bom_maps=bom_maps)
            else:
                result = ecn_to_pads_eco.run(ecn, lib, tmp, bom)
            skipped, pkg_changes = _eco_run(result)
    except SystemExit as e:  # 生成器的硬停止（BOM 无 Part Type 等）
        _log_gen_stdout(log, buf.getvalue())
        log(f"停止: {e}")
        _discard_generated(tmp, log)
        return False
    except EncryptedXlsxError as e:
        _log_gen_stdout(log, buf.getvalue())
        plain = user_file_error(e)
        if plain:
            log(f"错误: {plain}")
        else:
            log(f"错误: {e}")
            log("这份文件需要本机 Excel 才能读取。请安装 Excel，或先解密后再试。")
        _discard_generated(tmp, log)
        return False
    except Exception as e:  # 文件打不开 / 不是 Excel 等
        _log_gen_stdout(log, buf.getvalue())
        plain = user_file_error(e)
        if plain:
            log(f"错误: {plain}")
        else:
            log(f"错误: 无法读取输入文件 — {type(e).__name__}: {e}")
            if bom_maps is None:
                log("请确认 ECN 与 BOM 都是可用 Excel 打开的 .xlsx/.xlsm。")
            else:
                log("请确认 ECN 是可用 Excel 打开的 .xlsx/.xlsm。")
        _discard_generated(tmp, log)
        return False
    _log_gen_stdout(log, buf.getvalue())
    log("")
    bad = ecn_to_pads_eco.abnormal_skips(skipped)
    if bad:
        # 命令行没有弹窗，把原因打出来后继续。界面由 confirm 询问。
        if confirm is None:
            log(ecn_to_pads_eco.format_skip_dialog(bad).rstrip())
            log("")
        elif not confirm(bad):
            log(ecn_to_pads_eco.format_skip_dialog(bad).rstrip())
            log("")
            _discard_generated(tmp, log)
            log("已取消。未生成 ECO。")
            return False
        else:
            log(f"异常跳过 {len(bad)} 个位号，已选择继续。明细在对照表。")
            log("")
    if pkg_changes:
        if confirm_pkg is None:
            log(ecn_to_pads_eco.format_pkg_dialog(pkg_changes, ecn).rstrip())
            log("")
        elif not confirm_pkg(pkg_changes):
            log(ecn_to_pads_eco.format_pkg_dialog(pkg_changes, ecn).rstrip())
            log("")
            _discard_generated(tmp, log)
            log("已取消。未生成 ECO。")
            return False
        else:
            log(f"封装变化 {len(pkg_changes)} 个位号，已选择继续。明细在对照表。")
            log("")
    ok = do_check(tmp, bom, lib, log, label=out.name, bom_maps=bom_maps)
    if not ok:
        _discard_generated(tmp, log)
        log("未生成 ECO。")
        return False
    tmp_map = tmp.with_suffix(".map.csv")
    final_map = out.with_suffix(".map.csv")
    try:
        os.replace(tmp, out)
        if tmp_map.exists():
            os.replace(tmp_map, final_map)
    except OSError as e:
        _discard_generated(tmp, log)
        log(f"错误: 无法写出 {out.name} — {e}")
        return False
    log("")
    log(f"完成。{out.name}  对照表 {final_map.name}")
    return True


def _dialog_max_height(parent) -> int:
    """弹窗不超过工作区，避免按钮被任务栏挡住。"""
    if sys.platform == "win32":
        import ctypes

        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long),
            ]

        area = RECT()
        if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(area), 0):
            height = area.bottom - area.top
            if height > 200:
                return height - 48
    return max(parent.winfo_screenheight() - 80, 480)


def _clip_round(win, radius: int = 14) -> None:
    """把无边框窗口裁成圆角。失败时保持直角。"""
    if sys.platform != "win32":
        return
    import ctypes
    win.update_idletasks()
    hwnd = int(win.winfo_id())
    w, h = win.winfo_width(), win.winfo_height()
    if w < 20 or h < 20:
        return
    gdi32 = ctypes.windll.gdi32
    user32 = ctypes.windll.user32
    rgn = gdi32.CreateRoundRectRgn(0, 0, w + 1, h + 1, radius * 2, radius * 2)
    user32.SetWindowRgn(hwnd, rgn, True)


def _install_modal_min(parent, card, shade, title_row, title: str):
    """无边框弹窗加「—」。最小化后进任务栏，点任务栏还原。不置顶，方便看 Excel。"""
    import tkinter as tk
    from tkinter import font as tkfont

    stub = tk.Toplevel(parent)
    stub.withdraw()
    stub.title(title)
    stub.geometry("320x48")
    stub.resizable(False, False)
    for ico in (exe_dir() / "pads-eco.ico", bundled_dir() / "tool" / "pads-eco.ico", HERE / "pads-eco.ico"):
        if ico.exists():
            try:
                stub.iconbitmap(str(ico))
            except tk.TclError:
                pass
            break

    state = {"min": False, "ignore_map": False}

    def minimize(_e=None) -> None:
        if state["min"] or not card.winfo_exists():
            return
        state["min"] = True
        state["ignore_map"] = True
        try:
            card.grab_release()
        except tk.TclError:
            pass
        try:
            shade.withdraw()
            card.withdraw()
        except tk.TclError:
            pass
        stub.deiconify()
        stub.iconify()
        stub.after(300, lambda: state.update(ignore_map=False))

    def restore(_e=None) -> None:
        if not state["min"] or not card.winfo_exists():
            return
        state["min"] = False
        try:
            stub.withdraw()
        except tk.TclError:
            pass
        try:
            shade.deiconify()
            card.deiconify()
            card.lift()
            card.grab_set()
            card.focus_force()
        except tk.TclError:
            pass

    def on_stub_map(_e=None) -> None:
        if state["ignore_map"] or not state["min"]:
            return
        try:
            if stub.state() == "iconic":
                return
        except tk.TclError:
            return
        restore()

    stub.bind("<Map>", on_stub_map)
    f_min = tkfont.Font(family="Microsoft YaHei UI", size=16)
    btn = tk.Label(
        title_row, text="–", font=f_min, fg=INK_2, bg=WHITE, cursor="hand2", padx=10,
    )
    btn.pack(side="right")
    btn.bind("<Button-1>", minimize)
    btn.bind("<Enter>", lambda _e: btn.configure(fg=INK))
    btn.bind("<Leave>", lambda _e: btn.configure(fg=INK_2))
    return minimize, stub


def _skip_modal_line(row) -> dict:
    """弹窗一行。不改跳过判定，只把已有字段摊成位号 / 料号 / 类型 / 短原因。"""
    def txt(value) -> str:
        return "" if value is None else str(value).strip()

    reason = str(row[-1]) if row else ""
    old_pn = txt(row[2]) if len(row) > 2 else ""
    old_type = txt(row[3]) if len(row) > 3 else ""
    new_pn = txt(row[4]) if len(row) > 4 else ""
    new_type = txt(row[5]) if len(row) > 5 else ""
    if reason.startswith("新料号未入库:"):
        pn = new_pn or reason.split(":", 1)[-1].strip()
        pt = new_type
        tag = "未入器件库"
    elif reason in ("BOM无Value", "Value含非ASCII", "已带@无需删除"):
        # 这几类把 Value 放在类型那一格，不能当成 Part Type。
        pn = old_pn or new_pn
        pt = old_type
        tag = ecn_to_pads_eco.other_problem(reason)
    else:
        pn = new_pn or old_pn
        pt = new_type or old_type
        tag = ecn_to_pads_eco.other_problem(reason)
    return {"ref": txt(row[0]) or "—", "pn": pn or "—", "pt": pt or "—", "reason": tag or "—"}


def ask_continue_skips(parent, rows) -> bool:
    """异常跳过模态框。继续返回 True，取消或关闭返回 False。"""
    import tkinter as tk
    from tkinter import font as tkfont
    from PIL import ImageTk

    lines = [_skip_modal_line(row) for row in rows]
    counts: dict[str, int] = {}
    count_order: list[str] = []
    for line in lines:
        tag = line["reason"]
        if tag not in counts:
            counts[tag] = 0
            count_order.append(tag)
        counts[tag] += 1
    total = len(lines)

    ui = "Microsoft YaHei UI"
    f_title = tkfont.Font(family=ui, size=18, weight="bold")
    f_sub = tkfont.Font(family=ui, size=12)
    f_th = tkfont.Font(family=ui, size=12)
    f_td = tkfont.Font(family=ui, size=13)
    f_td_b = tkfont.Font(family=ui, size=13, weight="bold")
    f_tag = tkfont.Font(family=ui, size=12)
    f_btn = tkfont.Font(family=ui, size=14)
    f_btn_b = tkfont.Font(family=ui, size=14, weight="bold")
    f_chip = tkfont.Font(family=ui, size=15)
    f_num = tkfont.Font(family=ui, size=13, weight="bold")
    chip_h, badge_h = 48, 30
    pad_l, gap, pad_r = 20, 10, 18
    head_bg = "#F7F8FA"
    tag_bg, tag_fg = "#FDE8E8", "#9F1239"
    accent = "#F0B429"
    visible_rows = 5
    row_h = 36

    choice = {"ok": False}
    extras: list = []
    shade = tk.Toplevel(parent)
    shade.overrideredirect(True)
    shade.configure(bg="#0B1220")
    try:
        shade.attributes("-alpha", 0.38)
    except tk.TclError:
        pass
    shade.geometry(
        f"{parent.winfo_width()}x{parent.winfo_height()}+{parent.winfo_rootx()}+{parent.winfo_rooty()}"
    )

    card = tk.Toplevel(parent)
    card.overrideredirect(True)
    card.configure(bg=WHITE, highlightthickness=1, highlightbackground="#E6E8EC")

    def finish(ok: bool) -> None:
        choice["ok"] = ok
        try:
            card.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        for win in (card, shade, *extras):
            try:
                win.destroy()
            except tk.TclError:
                pass

    stripe = tk.Frame(card, bg=accent, height=4)
    stripe.pack(fill="x")
    stripe.pack_propagate(False)

    head = tk.Frame(card, bg=WHITE, padx=24)
    head.pack(fill="x", pady=(18, 8))
    title_row = tk.Frame(head, bg=WHITE)
    title_row.pack(fill="x")
    mark = tk.Canvas(title_row, width=28, height=28, highlightthickness=0, bg=WHITE, bd=0)
    mark.pack(side="left")
    mark_img = ImageTk.PhotoImage(_aa_mark())
    mark.create_image(0, 0, image=mark_img, anchor="nw")
    mark._photo = mark_img
    tk.Label(
        title_row, text=f"有 {total} 个位号因异常被跳过",
        font=f_title, fg=INK, bg=WHITE, anchor="w",
    ).pack(side="left", padx=(10, 0))
    _min, stub = _install_modal_min(
        parent, card, shade, title_row, f"有 {total} 个位号因异常被跳过",
    )
    extras.append(stub)
    stub.protocol("WM_DELETE_WINDOW", lambda: finish(False))
    tk.Label(
        head, text="导入此 ECO 时下列位号不会被修改，其余仍正常生成。",
        font=f_sub, fg=INK_2, bg=WHITE, anchor="w",
    ).pack(fill="x", pady=(8, 0))

    ref_w, pn_w, reason_w = 100, 168, 156

    def fill_cols(host, bg, ref, pn, pt, reason, header=False):
        host.grid_columnconfigure(0, minsize=ref_w)
        host.grid_columnconfigure(1, minsize=pn_w)
        host.grid_columnconfigure(2, weight=1)
        host.grid_columnconfigure(3, minsize=reason_w)
        host.grid_rowconfigure(0, minsize=32 if header else row_h - 1)
        fg = INK_2 if header else INK
        tk.Label(host, text=ref, font=f_th if header else f_td_b, fg=fg, bg=bg, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(12, 8),
        )
        tk.Label(host, text=pn, font=f_th if header else f_td, fg=fg, bg=bg, anchor="w").grid(
            row=0, column=1, sticky="w", padx=(0, 8),
        )
        tk.Label(host, text=pt, font=f_th if header else f_td, fg=fg, bg=bg, anchor="w").grid(
            row=0, column=2, sticky="we", padx=(0, 8),
        )
        if header:
            tk.Label(host, text=reason, font=f_th, fg=INK_2, bg=bg, anchor="w").grid(
                row=0, column=3, sticky="w", padx=(0, 12),
            )
        else:
            tk.Label(host, text=reason, font=f_tag, fg=tag_fg, bg=tag_bg, padx=8, pady=1).grid(
                row=0, column=3, sticky="w", padx=(0, 12),
            )

    selected = {"tag": count_order[0] if count_order else ""}
    body = tk.Frame(card, bg=WHITE, padx=24)
    body.pack(fill="x", pady=(8, 0))
    shell = tk.Frame(body, bg=WHITE, highlightthickness=1, highlightbackground=LINE)
    shell.pack(fill="x")
    header = tk.Frame(shell, bg=head_bg)
    header.pack(fill="x")
    spacer = tk.Frame(header, bg=head_bg, width=14)
    fill_cols(header, head_bg, "位号", "Part Number", "Part Type", "异常原因", header=True)
    tk.Frame(shell, bg=LINE, height=1).pack(fill="x")

    from tkinter import ttk as _ttk
    body_wrap = tk.Frame(shell, bg=WHITE)
    body_wrap.pack(fill="x")
    canvas = tk.Canvas(
        body_wrap, bg=WHITE, height=visible_rows * row_h, highlightthickness=0, bd=0,
    )
    style = _ttk.Style(card)
    style.configure(
        "Skip.Vertical.TScrollbar",
        background="#D5DAE1", troughcolor=WHITE, bordercolor=WHITE, arrowcolor="#6B7280",
        gripcount=0, borderwidth=0, lightcolor="#D5DAE1", darkcolor="#D5DAE1",
    )
    vsb = _ttk.Scrollbar(body_wrap, orient="vertical", command=canvas.yview, style="Skip.Vertical.TScrollbar")
    canvas.configure(yscrollcommand=vsb.set)
    canvas.pack(side="left", fill="both", expand=True)
    inner = tk.Frame(canvas, bg=WHITE)
    wid = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _fit_table(event):
        canvas.itemconfigure(wid, width=event.width)

    canvas.bind("<Configure>", _fit_table)

    def _wheel(event):
        canvas.yview_scroll(int(-event.delta / 120), "units")

    card.bind_all("<MouseWheel>", _wheel)

    summary = tk.Frame(card, bg=WHITE, padx=24)
    summary.pack(fill="x", pady=(16, 4))
    flow = tk.Frame(summary, bg=WHITE)
    flow.pack(fill="x")
    line_box = None
    used = 0
    limit = 800 - 48
    chips = []

    def paint_chip(box, text, count, hot):
        num = str(count)
        badge = max(badge_h, f_num.measure(num) + 16)
        width = int(box.cget("width"))
        height = int(box.cget("height"))
        fill = "#FFF7ED" if hot else "#F5F5F5"
        outline = "#EA580C" if hot else "#E5E5E5"
        fg = "#9A3412" if hot else "#525252"
        badge_fill = "#EA580C" if hot else "#737373"
        pill = ImageTk.PhotoImage(_aa_rounded(width, height, height / 2, fill, outline, 2 if hot else 1))
        blob = ImageTk.PhotoImage(_aa_rounded(badge, badge_h, badge_h / 2, badge_fill, bg=fill))
        box.delete("all")
        box.create_image(0, 0, image=pill, anchor="nw")
        cy = height / 2
        cx = pad_l + f_chip.measure(text) + gap + badge / 2
        box.create_image(cx, cy, image=blob, anchor="center")
        box.create_text(pad_l, cy, text=text, anchor="w", fill=fg, font=f_chip)
        box.create_text(cx, cy, text=num, fill=WHITE, font=f_num)
        box._photos = (pill, blob)

    def show_reason(tag):
        selected["tag"] = tag
        for child in inner.winfo_children():
            child.destroy()
        view = [line for line in lines if line["reason"] == tag]
        for line in view:
            row = tk.Frame(inner, bg=WHITE, height=row_h - 1)
            row.pack(fill="x")
            row.pack_propagate(False)
            fill_cols(
                row, WHITE,
                _ellipsize(f_td_b, line["ref"], ref_w - 24),
                _ellipsize(f_td, line["pn"], pn_w - 12),
                _ellipsize(f_td, line["pt"], 300),
                line["reason"],
            )
            tk.Frame(inner, bg="#F3F4F6", height=1).pack(fill="x")
        inner.update_idletasks()
        canvas.itemconfigure(wid, width=max(canvas.winfo_width(), 1))
        canvas.configure(scrollregion=canvas.bbox("all") or (0, 0, 1, 1))
        canvas.yview_moveto(0)
        show_bar = len(view) > visible_rows
        header.grid_columnconfigure(4, minsize=14 if show_bar else 0)
        if show_bar:
            spacer.grid(row=0, column=4, sticky="ns")
            vsb.pack(side="right", fill="y", before=canvas)
        else:
            spacer.grid_remove()
            vsb.pack_forget()
        for chip in chips:
            paint_chip(chip["box"], chip["tag"], chip["count"], chip["tag"] == tag)

    for tag in count_order:
        num = str(counts[tag])
        badge = max(badge_h, f_num.measure(num) + 16)
        tw = pad_l + f_chip.measure(tag) + gap + badge + pad_r
        if line_box is None or used + tw > limit:
            line_box = tk.Frame(flow, bg=WHITE)
            line_box.pack(fill="x", pady=(0, 8))
            used = 0
        box = tk.Canvas(line_box, width=tw, height=chip_h, highlightthickness=0, bg=WHITE, bd=0, cursor="hand2")
        box.pack(side="left", padx=(0, 8))
        box.bind("<Button-1>", lambda _e, picked=tag: show_reason(picked))
        chips.append({"box": box, "tag": tag, "count": counts[tag]})
        used += tw + 8

    show_reason(selected["tag"])

    tk.Frame(card, bg=LINE, height=1).pack(fill="x", pady=(8, 0))
    bar = tk.Frame(card, bg=WHITE, padx=24, pady=16)
    bar.pack(fill="x")
    go = tk.Label(bar, text="继续生成", font=f_btn_b, fg=WHITE, bg=GREEN, padx=22, pady=8, cursor="hand2")
    stop = tk.Label(
        bar, text="取消", font=f_btn, fg=INK, bg=WHITE, padx=22, pady=8, cursor="hand2",
        highlightthickness=1, highlightbackground=LINE,
    )
    go.pack(side="right")
    stop.pack(side="right", padx=(0, 8))
    go.bind("<Button-1>", lambda _e: finish(True))
    go.bind("<Enter>", lambda _e: go.configure(bg=GREEN_DK))
    go.bind("<Leave>", lambda _e: go.configure(bg=GREEN))
    stop.bind("<Button-1>", lambda _e: finish(False))
    stop.bind("<Enter>", lambda _e: stop.configure(bg=head_bg))
    stop.bind("<Leave>", lambda _e: stop.configure(bg=WHITE))
    card.bind("<Escape>", lambda _e: finish(False))
    shade.bind("<Button-1>", lambda _e: finish(False))

    card.update_idletasks()
    w = 800
    h = max(card.winfo_reqheight(), 200)
    x = parent.winfo_rootx() + max(0, (parent.winfo_width() - w) // 2)
    y = parent.winfo_rooty() + max(0, (parent.winfo_height() - h) // 2)
    if sys.platform == "win32":
        import ctypes
        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long),
            ]
        area = RECT()
        if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(area), 0):
            x = max(area.left + 12, min(x, area.right - w - 12))
            y = max(area.top + 12, min(y, area.bottom - h - 12))
    card.geometry(f"{w}x{h}+{x}+{y}")
    _clip_round(card, 14)
    card.grab_set()
    card.lift()
    card.focus_force()
    parent.wait_window(card)
    try:
        if shade.winfo_exists():
            shade.destroy()
    except tk.TclError:
        pass
    return choice["ok"]


def _pkg_modal_line(row) -> dict:
    """弹窗一行：位号 / 变更类型 / 变更前封装 / 变更后封装。"""
    def txt(value) -> str:
        return "" if value is None else str(value).strip()

    before = txt(row[2]) if len(row) > 2 else ""
    after = txt(row[3]) if len(row) > 3 else ""
    return {
        "ref": txt(row[0]) or "—",
        "action": txt(row[1]) if len(row) > 1 else "—",
        "before": before or "—",
        "after": after or "—",
        "pair": f"{before or '—'} → {after or '—'}",
    }


def _open_ecn_file(path) -> None:
    """用系统默认程序打开这次选中的 ECN。失败时静默，不打断弹窗。"""
    if not path:
        return
    p = Path(path)
    if not p.exists():
        return
    try:
        os.startfile(str(p))
    except OSError:
        pass


def ask_continue_pkg(parent, rows, ecn=None) -> bool:
    """封装变化模态框。继续返回 True，取消或关闭返回 False。"""
    import tkinter as tk
    from tkinter import font as tkfont
    from PIL import ImageTk

    lines = [_pkg_modal_line(row) for row in rows]
    counts: dict[str, int] = {}
    count_order: list[str] = []
    for line in lines:
        tag = line["pair"]
        if tag not in counts:
            counts[tag] = 0
            count_order.append(tag)
        counts[tag] += 1
    total = len(lines)

    ui = "Microsoft YaHei UI"
    f_title = tkfont.Font(family=ui, size=18, weight="bold")
    f_sub = tkfont.Font(family=ui, size=12)
    f_th = tkfont.Font(family=ui, size=12)
    f_td = tkfont.Font(family=ui, size=13)
    f_td_b = tkfont.Font(family=ui, size=13, weight="bold")
    f_tag = tkfont.Font(family=ui, size=12)
    f_btn = tkfont.Font(family=ui, size=14)
    f_btn_b = tkfont.Font(family=ui, size=14, weight="bold")
    f_chip = tkfont.Font(family=ui, size=15)
    f_num = tkfont.Font(family=ui, size=13, weight="bold")
    chip_h, badge_h = 48, 30
    pad_l, gap, pad_r = 20, 10, 18
    head_bg = "#F7F8FA"
    tag_bg, tag_fg = "#FFF7ED", "#C05621"
    accent = "#F0B429"
    visible_rows = 5
    row_h = 36

    choice = {"ok": False}
    extras: list = []
    shade = tk.Toplevel(parent)
    shade.overrideredirect(True)
    shade.configure(bg="#0B1220")
    try:
        shade.attributes("-alpha", 0.38)
    except tk.TclError:
        pass
    shade.geometry(
        f"{parent.winfo_width()}x{parent.winfo_height()}+{parent.winfo_rootx()}+{parent.winfo_rooty()}"
    )

    card = tk.Toplevel(parent)
    card.overrideredirect(True)
    card.configure(bg=WHITE, highlightthickness=1, highlightbackground="#E6E8EC")

    def finish(ok: bool) -> None:
        choice["ok"] = ok
        try:
            card.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        for win in (card, shade, *extras):
            try:
                win.destroy()
            except tk.TclError:
                pass

    stripe = tk.Frame(card, bg=accent, height=4)
    stripe.pack(fill="x")
    stripe.pack_propagate(False)

    head = tk.Frame(card, bg=WHITE, padx=24)
    head.pack(fill="x", pady=(18, 8))
    title_row = tk.Frame(head, bg=WHITE)
    title_row.pack(fill="x")
    mark = tk.Canvas(title_row, width=28, height=28, highlightthickness=0, bg=WHITE, bd=0)
    mark.pack(side="left")
    mark_img = ImageTk.PhotoImage(_aa_mark())
    mark.create_image(0, 0, image=mark_img, anchor="nw")
    mark._photo = mark_img
    tk.Label(
        title_row, text=f"有 {total} 个位号封装发生变化",
        font=f_title, fg=INK, bg=WHITE, anchor="w",
    ).pack(side="left", padx=(10, 0))
    minimize, stub = _install_modal_min(
        parent, card, shade, title_row, f"有 {total} 个位号封装发生变化",
    )
    extras.append(stub)
    stub.protocol("WM_DELETE_WINDOW", lambda: finish(False))
    hint = tk.Text(
        head, font=f_sub, fg=INK_2, bg=WHITE, bd=0, highlightthickness=0,
        wrap="word", height=2, cursor="arrow", padx=0, pady=0, relief="flat",
        takefocus=0, exportselection=False,
    )
    hint.insert("1.0", "请先打开 ")
    hint.insert("end", "ECN 变更表", "link")
    hint.insert("end", "，核对下列封装变化是否正确。确认无误后再继续生成。")
    hint.tag_configure("link", foreground="#2563EB", underline=True)
    hint.tag_bind("link", "<Enter>", lambda _e: hint.configure(cursor="hand2"))
    hint.tag_bind("link", "<Leave>", lambda _e: hint.configure(cursor="arrow"))

    def open_ecn(_e=None):
        minimize()
        _open_ecn_file(ecn)

    hint.tag_bind("link", "<Button-1>", open_ecn)
    hint.bind("<Key>", lambda _e: "break")
    hint.pack(fill="x", pady=(8, 0))

    ref_w, action_w, before_w, after_w = 100, 120, 140, 140

    def fill_cols(host, bg, ref, action, before, after, header=False):
        host.grid_columnconfigure(0, minsize=ref_w)
        host.grid_columnconfigure(1, minsize=action_w)
        host.grid_columnconfigure(2, minsize=before_w)
        host.grid_columnconfigure(3, weight=1, minsize=after_w)
        host.grid_rowconfigure(0, minsize=32 if header else row_h - 1)
        fg = INK_2 if header else INK
        tk.Label(host, text=ref, font=f_th if header else f_td_b, fg=fg, bg=bg, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(12, 8),
        )
        tk.Label(host, text=action, font=f_th if header else f_td, fg=fg, bg=bg, anchor="w").grid(
            row=0, column=1, sticky="w", padx=(0, 8),
        )
        tk.Label(host, text=before, font=f_th if header else f_td, fg=fg, bg=bg, anchor="w").grid(
            row=0, column=2, sticky="w", padx=(0, 8),
        )
        if header:
            tk.Label(host, text=after, font=f_th, fg=INK_2, bg=bg, anchor="w").grid(
                row=0, column=3, sticky="w", padx=(0, 12),
            )
        else:
            tk.Label(host, text=after, font=f_tag, fg=tag_fg, bg=tag_bg, padx=8, pady=1).grid(
                row=0, column=3, sticky="w", padx=(0, 12),
            )

    selected = {"tag": count_order[0] if count_order else ""}
    body = tk.Frame(card, bg=WHITE, padx=24)
    body.pack(fill="x", pady=(8, 0))
    shell = tk.Frame(body, bg=WHITE, highlightthickness=1, highlightbackground=LINE)
    shell.pack(fill="x")
    header = tk.Frame(shell, bg=head_bg)
    header.pack(fill="x")
    spacer = tk.Frame(header, bg=head_bg, width=14)
    fill_cols(header, head_bg, "位号", "变更类型", "变更前封装", "变更后封装", header=True)
    tk.Frame(shell, bg=LINE, height=1).pack(fill="x")

    from tkinter import ttk as _ttk
    body_wrap = tk.Frame(shell, bg=WHITE)
    body_wrap.pack(fill="x")
    canvas = tk.Canvas(
        body_wrap, bg=WHITE, height=visible_rows * row_h, highlightthickness=0, bd=0,
    )
    style = _ttk.Style(card)
    style.configure(
        "Pkg.Vertical.TScrollbar",
        background="#D5DAE1", troughcolor=WHITE, bordercolor=WHITE, arrowcolor="#6B7280",
        gripcount=0, borderwidth=0, lightcolor="#D5DAE1", darkcolor="#D5DAE1",
    )
    vsb = _ttk.Scrollbar(body_wrap, orient="vertical", command=canvas.yview, style="Pkg.Vertical.TScrollbar")
    canvas.configure(yscrollcommand=vsb.set)
    canvas.pack(side="left", fill="both", expand=True)
    inner = tk.Frame(canvas, bg=WHITE)
    wid = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _fit_table(event):
        canvas.itemconfigure(wid, width=event.width)

    canvas.bind("<Configure>", _fit_table)

    def _wheel(event):
        canvas.yview_scroll(int(-event.delta / 120), "units")

    card.bind_all("<MouseWheel>", _wheel)

    summary = tk.Frame(card, bg=WHITE, padx=24)
    summary.pack(fill="x", pady=(16, 4))
    flow = tk.Frame(summary, bg=WHITE)
    flow.pack(fill="x")
    line_box = None
    used = 0
    limit = 800 - 48
    chips = []

    def paint_chip(box, text, count, hot):
        num = str(count)
        badge = max(badge_h, f_num.measure(num) + 16)
        width = int(box.cget("width"))
        height = int(box.cget("height"))
        fill = "#FFF7ED" if hot else "#F5F5F5"
        outline = "#F5C48A" if hot else "#E5E5E5"
        fg = "#C05621" if hot else "#525252"
        badge_fill = "#C05621" if hot else "#737373"
        pill = ImageTk.PhotoImage(_aa_rounded(width, height, height / 2, fill, outline, 2 if hot else 1))
        blob = ImageTk.PhotoImage(_aa_rounded(badge, badge_h, badge_h / 2, badge_fill, bg=fill))
        box.delete("all")
        box.create_image(0, 0, image=pill, anchor="nw")
        cy = height / 2
        cx = pad_l + f_chip.measure(text) + gap + badge / 2
        box.create_image(cx, cy, image=blob, anchor="center")
        box.create_text(pad_l, cy, text=text, anchor="w", fill=fg, font=f_chip)
        box.create_text(cx, cy, text=num, fill=WHITE, font=f_num)
        box._photos = (pill, blob)

    def show_pair(tag):
        selected["tag"] = tag
        for child in inner.winfo_children():
            child.destroy()
        view = [line for line in lines if line["pair"] == tag]
        for line in view:
            row = tk.Frame(inner, bg=WHITE, height=row_h - 1)
            row.pack(fill="x")
            row.pack_propagate(False)
            fill_cols(
                row, WHITE,
                _ellipsize(f_td_b, line["ref"], ref_w - 24),
                _ellipsize(f_td, line["action"], action_w - 12),
                _ellipsize(f_td, line["before"], before_w - 12),
                line["after"],
            )
            tk.Frame(inner, bg="#F3F4F6", height=1).pack(fill="x")
        inner.update_idletasks()
        canvas.itemconfigure(wid, width=max(canvas.winfo_width(), 1))
        canvas.configure(scrollregion=canvas.bbox("all") or (0, 0, 1, 1))
        canvas.yview_moveto(0)
        show_bar = len(view) > visible_rows
        header.grid_columnconfigure(4, minsize=14 if show_bar else 0)
        if show_bar:
            spacer.grid(row=0, column=4, sticky="ns")
            vsb.pack(side="right", fill="y", before=canvas)
        else:
            spacer.grid_remove()
            vsb.pack_forget()
        for chip in chips:
            paint_chip(chip["box"], chip["tag"], chip["count"], chip["tag"] == tag)

    for tag in count_order:
        num = str(counts[tag])
        badge = max(badge_h, f_num.measure(num) + 16)
        tw = pad_l + f_chip.measure(tag) + gap + badge + pad_r
        if line_box is None or used + tw > limit:
            line_box = tk.Frame(flow, bg=WHITE)
            line_box.pack(fill="x", pady=(0, 8))
            used = 0
        box = tk.Canvas(line_box, width=tw, height=chip_h, highlightthickness=0, bg=WHITE, bd=0, cursor="hand2")
        box.pack(side="left", padx=(0, 8))
        box.bind("<Button-1>", lambda _e, picked=tag: show_pair(picked))
        chips.append({"box": box, "tag": tag, "count": counts[tag]})
        used += tw + 8

    show_pair(selected["tag"])

    tk.Frame(card, bg=LINE, height=1).pack(fill="x", pady=(8, 0))
    bar = tk.Frame(card, bg=WHITE, padx=24, pady=16)
    bar.pack(fill="x")
    go = tk.Label(bar, text="继续生成", font=f_btn_b, fg=WHITE, bg=GREEN, padx=22, pady=8, cursor="hand2")
    stop = tk.Label(
        bar, text="取消", font=f_btn, fg=INK, bg=WHITE, padx=22, pady=8, cursor="hand2",
        highlightthickness=1, highlightbackground=LINE,
    )
    go.pack(side="right")
    stop.pack(side="right", padx=(0, 8))
    go.bind("<Button-1>", lambda _e: finish(True))
    go.bind("<Enter>", lambda _e: go.configure(bg=GREEN_DK))
    go.bind("<Leave>", lambda _e: go.configure(bg=GREEN))
    stop.bind("<Button-1>", lambda _e: finish(False))
    stop.bind("<Enter>", lambda _e: stop.configure(bg=head_bg))
    stop.bind("<Leave>", lambda _e: stop.configure(bg=WHITE))
    card.bind("<Escape>", lambda _e: finish(False))
    shade.bind("<Button-1>", lambda _e: finish(False))

    card.update_idletasks()
    w = 800
    h = max(card.winfo_reqheight(), 200)
    x = parent.winfo_rootx() + max(0, (parent.winfo_width() - w) // 2)
    y = parent.winfo_rooty() + max(0, (parent.winfo_height() - h) // 2)
    if sys.platform == "win32":
        import ctypes
        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long),
            ]
        area = RECT()
        if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(area), 0):
            x = max(area.left + 12, min(x, area.right - w - 12))
            y = max(area.top + 12, min(y, area.bottom - h - 12))
    card.geometry(f"{w}x{h}+{x}+{y}")
    _clip_round(card, 14)
    card.grab_set()
    card.lift()
    card.focus_force()
    parent.wait_window(card)
    try:
        if shade.winfo_exists():
            shade.destroy()
    except tk.TclError:
        pass
    return choice["ok"]


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def cli(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(prog="pads-eco", description=f"PADS Logic ECO 生成/校验工具 v{VERSION}")
    sub = ap.add_subparsers(dest="cmd", required=True)

    g = sub.add_parser("gen", help="从 ECN + 原理图 BOM 生成 ECO 并校验")
    g.add_argument("ecn", type=Path)
    g.add_argument("--bom", required=True, type=Path, help="PADS Logic 导出的带 Part Type 的 BOM")
    g.add_argument("--lib", type=Path, default=None, help=f"器件库 xlsx（默认 exe 旁 library\\{LIB_NAME} 或内置）")
    g.add_argument("-o", "--output", type=Path, default=None)

    c = sub.add_parser("check", help="校验已有 ECO")
    c.add_argument("eco", type=Path)
    c.add_argument("--bom", required=True, type=Path)
    c.add_argument("--lib", type=Path, default=None)

    a = ap.parse_args(argv)
    lib = a.lib or default_lib()
    if not lib.exists():
        print(f"找不到器件库: {lib}")
        return 2
    log = print
    if a.cmd == "gen":
        return 0 if do_gen(a.ecn, a.bom, lib, a.output, log) else 1
    return 0 if do_check(a.eco, a.bom, lib, log) else 1


# ---------------------------------------------------------------------------
# GUI — 浅灰白底，青绿 #0F6E56。左侧圆点竖线是五步顺序。
# ---------------------------------------------------------------------------
BG = "#F3F4F6"
BAR_H = 30
BAR_LINE = "#E5E7EB"
TITLE_FG = "#111827"
INK = "#1C1917"
INK_2 = "#5F6772"
INK_3 = "#8B939E"
GREEN = "#0F6E56"
GREEN_DK = "#0C5946"
GREEN_SOFT = "#E7F5F0"
GREEN_LINE = "#C5E4D8"
AMBER = "#C05621"
AMBER_BG = "#FFF7ED"
AMBER_LINE = "#F5C48A"
LOG_BG = "#070B12"
LOG_FG = "#FFFFFF"
LOG_OK = "#9AF6C4"
LOG_BAD = "#FFB4B4"
LOG_HINT = "#FFFFFF"
LINE = "#E6E8EC"
DASH = "#C5CCD4"
CUR_LINE = "#93C5FD"
CUR_RING = "#E7F0FE"
CUR_DOT = "#2563EB"
FAIL_C = "#C2410C"
WHITE = "#FFFFFF"

UI = "Microsoft YaHei UI"
MONO = "Consolas"


def _round_rect(c, x1, y1, x2, y2, r, **kw):
    """Crisp corners: arcs and rectangles, not a smoothed polygon."""
    tags = kw.pop("tags", ())
    fill = kw.pop("fill", "")
    outline = kw.pop("outline", "")
    width = kw.pop("width", 1) or 1
    dash = kw.pop("dash", None)
    tagkw = {"tags": tags} if tags else {}
    x1, y1, x2, y2 = float(x1), float(y1), float(x2), float(y2)
    r = max(0.0, min(float(r), (x2 - x1) / 2.0, (y2 - y1) / 2.0))
    if fill:
        if r >= 1:
            c.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline="", **tagkw)
            c.create_rectangle(x1, y1 + r, x2, y2 - r, fill=fill, outline="", **tagkw)
            for box, start in (
                ((x1, y1, x1 + 2 * r, y1 + 2 * r), 90),
                ((x2 - 2 * r, y1, x2, y1 + 2 * r), 0),
                ((x2 - 2 * r, y2 - 2 * r, x2, y2), 270),
                ((x1, y2 - 2 * r, x1 + 2 * r, y2), 180),
            ):
                c.create_arc(*box, start=start, extent=90, fill=fill, outline="", style="pieslice", **tagkw)
        else:
            c.create_rectangle(x1, y1, x2, y2, fill=fill, outline="", **tagkw)
    if outline:
        _stroke_round(c, x1, y1, x2, y2, r, outline, width, dash)


def _stroke_round(c, x1, y1, x2, y2, r, outline, width=1, dash=None):
    kw = {"style": "arc", "outline": outline, "width": width}
    if dash:
        kw["dash"] = dash
    c.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, **kw)
    c.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, **kw)
    c.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, **kw)
    c.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, **kw)
    line = {"fill": outline, "width": width}
    if dash:
        line["dash"] = dash
    c.create_line(x1 + r, y1, x2 - r, y1, **line)
    c.create_line(x2, y1 + r, x2, y2 - r, **line)
    c.create_line(x2 - r, y2, x1 + r, y2, **line)
    c.create_line(x1, y2 - r, x1, y1 + r, **line)


def _dot_rgba(kind: str):
    """Anti-aliased stepper dot. Canvas ovals are stair-stepped at this size."""
    from PIL import Image, ImageDraw

    scale = 4
    diameter = 22
    pad = 4
    size = diameter + pad * 2
    img = Image.new("RGBA", (size * scale, size * scale), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    m = pad * scale
    box = [m, m, m + diameter * scale - 1, m + diameter * scale - 1]
    if kind == "done":
        draw.ellipse(box, fill=GREEN)
        cx = (box[0] + box[2]) / 2
        cy = (box[1] + box[3]) / 2 + scale * 0.4
        draw.line(
            [(cx - 5 * scale, cy - 0.2 * scale), (cx - 1.6 * scale, cy + 3.4 * scale), (cx + 5.4 * scale, cy - 4.2 * scale)],
            fill=WHITE,
            width=2 * scale,
        )
    elif kind == "current":
        # 圆点外沿之外 3px 的淡蓝环，描边落在 22px 圆内，不改步进占位。
        halo = 3 * scale
        hbox = [box[0] - halo, box[1] - halo, box[2] + halo, box[3] + halo]
        draw.ellipse(hbox, fill=(37, 99, 235, 38))
        draw.ellipse(box, fill=WHITE)
        stroke = 2 * scale
        inset = stroke / 2
        ibox = [box[0] + inset, box[1] + inset, box[2] - inset, box[3] - inset]
        draw.ellipse(ibox, outline=CUR_DOT, width=stroke)
    elif kind == "gen_on":
        draw.ellipse(box, fill=WHITE, outline=GREEN, width=2 * scale)
    else:
        draw.ellipse(box, fill=WHITE, outline=DASH, width=scale)
    return img.resize((size, size), Image.Resampling.LANCZOS)


def _hex_rgb(color: str) -> tuple[int, int, int]:
    color = color.lstrip("#")
    return (int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16))


def _aa_rounded(width: int, height: int, radius: float, fill: str, outline: str = "", stroke: float = 1, bg: str = WHITE):
    """圆角图，先放大再缩回，并铺在底色上。画布圆弧没有抗锯齿，边缘会是一格一格的。"""
    from PIL import Image, ImageDraw

    scale = 4
    img = Image.new("RGBA", (width * scale, height * scale), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    pad = (stroke / 2 + 0.75) if outline else 0.75
    box = [pad * scale, pad * scale, (width - pad) * scale - 1, (height - pad) * scale - 1]
    if box[2] > box[0] and box[3] > box[1]:
        rad = max(0.0, min(float(radius) * scale, (box[2] - box[0]) / 2, (box[3] - box[1]) / 2))
        kw = {"radius": rad, "fill": (*_hex_rgb(fill), 255)}
        if outline:
            kw["outline"] = (*_hex_rgb(outline), 255)
            kw["width"] = max(1, int(round(stroke * scale)))
        draw.rounded_rectangle(box, **kw)
    smooth = img.resize((width, height), Image.Resampling.LANCZOS)
    base = Image.new("RGBA", (width, height), (*_hex_rgb(bg), 255))
    base.alpha_composite(smooth)
    return base.convert("RGB")


def _aa_mark():
    """警告圆和感叹号画在同一张图上，避免圆的锯齿和文字的彩色边。"""
    from PIL import Image, ImageDraw, ImageFont

    scale = 4
    size = 28
    img = Image.new("RGBA", (size * scale, size * scale), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    pad = 1.25 * scale
    box = [pad, pad, size * scale - 1 - pad, size * scale - 1 - pad]
    draw.rounded_rectangle(
        box, radius=(box[2] - box[0]) / 2,
        fill=(*_hex_rgb("#FFF8EB"), 255), outline=(*_hex_rgb("#F0B429"), 255), width=scale,
    )
    try:
        font = ImageFont.truetype(r"C:\Windows\Fonts\msyhbd.ttc", 15 * scale, index=1)
    except OSError:
        font = ImageFont.load_default()
    draw.text((size * scale / 2, size * scale / 2 - scale * 0.4), "!", font=font, fill=(*_hex_rgb(AMBER), 255), anchor="mm")
    smooth = img.resize((size, size), Image.Resampling.LANCZOS)
    base = Image.new("RGBA", (size, size), (*_hex_rgb(WHITE), 255))
    base.alpha_composite(smooth)
    return base.convert("RGB")


def _ellipsize(fnt, text, width):
    if width <= 0 or fnt.measure(text) <= width:
        return text
    ell = "…"
    lo, hi = 0, len(text)
    while lo < hi:
        mid = (lo + hi) // 2
        if fnt.measure(text[:mid] + ell) <= width:
            lo = mid + 1
        else:
            hi = mid
    return text[: max(0, lo - 1)] + ell


def _wrap_text(fnt, text, width):
    lines, cur = [], ""
    for ch in text:
        if cur and fnt.measure(cur + ch) > width:
            lines.append(cur)
            cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    return lines or [""]


def gui() -> int:
    import re
    import tkinter as tk
    from tkinter import filedialog, font as tkfont, ttk

    root = tk.Tk()
    root.title("PADS Logic ECO")
    root.configure(bg=BG)
    # 1 像素 = 1 点，避免系统再把低分辨率画面放大发虚。宽度按 1080p 留出文件名。
    try:
        root.tk.call("tk", "scaling", 1.0)
    except tk.TclError:
        pass
    root.geometry("1000x890+80+30")
    root.minsize(880, 750)
    for ico in (exe_dir() / "pads-eco.ico", bundled_dir() / "tool" / "pads-eco.ico", HERE / "pads-eco.ico"):
        if ico.exists():
            try:
                root.iconbitmap(default=str(ico))
            except tk.TclError:
                pass
            break

    fams = set(tkfont.families())
    ui = UI if UI in fams else "Microsoft YaHei" if "Microsoft YaHei" in fams else "Segoe UI"
    mono = MONO if MONO in fams else "Courier New"
    f_card = tkfont.Font(root=root, family=ui, size=15, weight="bold")
    f_file = tkfont.Font(root=root, family=ui, size=13)
    f_badge = tkfont.Font(root=root, family=ui, size=11, weight="bold")
    f_ghost = tkfont.Font(root=root, family=ui, size=13)
    f_gen = tkfont.Font(root=root, family=ui, size=16, weight="bold")
    f_num = tkfont.Font(root=root, family=mono, size=11, weight="bold")
    f_warn = tkfont.Font(root=root, family=ui, size=11)
    f_run = tkfont.Font(root=root, family=ui, size=11)
    f_run_ok = tkfont.Font(root=root, family=ui, size=14, weight="bold")
    f_result = tkfont.Font(root=root, family=ui, size=20, weight="bold")
    f_stat = tkfont.Font(root=root, family=ui, size=18, weight="bold")
    f_guide_h = tkfont.Font(root=root, family=ui, size=16, weight="bold")
    title_face = "Segoe UI" if "Segoe UI" in fams else ui
    f_title = tkfont.Font(root=root, family=title_face, size=12, weight="bold")

    state = {
        "ecn": None, "bom": None, "lib": default_lib(), "last_out": None, "lib_ok": False,
        "lib_step_done": False,
        "gen_ok": False, "rev_ok": False, "rev_name": None, "rev_book": None,
        "ready": False, "busy": False,
        "hover": None, "press": None, "pressed": False, "painting": False, "hits": [], "view": None,
        "dot_img": {},
    }

    chrome = {"zoom": False, "geom": None, "drag": None}

    bar = tk.Frame(root, bg=WHITE, height=BAR_H, highlightthickness=0, bd=0)
    bar.pack(side="top", fill="x")
    bar.pack_propagate(False)
    tk.Frame(bar, bg=BAR_LINE, height=1).pack(side="bottom", fill="x")
    title_lbl = tk.Label(bar, text="PADS Logic ECO", font=f_title, fg=TITLE_FG, bg=WHITE)
    title_lbl.pack(side="left", padx=(12, 0))

    def paint_caption(c, kind, bg, fg):
        c.delete("all")
        c.configure(bg=bg)
        cx, cy = 23, 14
        if kind == "min":
            c.create_line(cx - 5, cy, cx + 5, cy, fill=fg, width=1)
        elif kind == "max":
            if chrome["zoom"]:
                c.create_rectangle(cx - 5, cy - 1, cx + 2, cy + 5, outline=fg, width=1)
                c.create_line(cx - 2, cy - 4, cx + 5, cy - 4, fill=fg, width=1)
                c.create_line(cx + 5, cy - 4, cx + 5, cy + 2, fill=fg, width=1)
                c.create_line(cx - 2, cy - 4, cx - 2, cy - 1, fill=fg, width=1)
                c.create_line(cx - 2, cy + 2, cx + 5, cy + 2, fill=fg, width=1)
            else:
                c.create_rectangle(cx - 5, cy - 4, cx + 5, cy + 5, outline=fg, width=1)
        else:
            c.create_line(cx - 4, cy - 4, cx + 4, cy + 4, fill=fg, width=1)
            c.create_line(cx + 4, cy - 4, cx - 4, cy + 4, fill=fg, width=1)

    caption_paint = {}

    def caption_button(kind):
        c = tk.Canvas(bar, width=46, height=BAR_H - 1, bg=WHITE, highlightthickness=0, bd=0, cursor="arrow")
        c.pack(side="right")
        paint_caption(c, kind, WHITE, TITLE_FG)
        caption_paint[kind] = c

        def enter(_e):
            if kind == "close":
                paint_caption(c, kind, "#E81123", WHITE)
            else:
                paint_caption(c, kind, "#F3F4F6", TITLE_FG)

        def leave(_e):
            paint_caption(c, kind, WHITE, TITLE_FG)

        c.bind("<Enter>", enter)
        c.bind("<Leave>", leave)
        return c

    btn_close = caption_button("close")
    btn_max = caption_button("max")
    btn_min = caption_button("min")

    col = tk.Frame(root, bg=BG, padx=18, pady=14)
    col.pack(fill="both", expand=True)

    stepper = tk.Canvas(col, bg=BG, height=380, highlightthickness=0, bd=0)
    stepper.pack(fill="x")
    warn_cv = tk.Canvas(col, bg=BG, height=52, highlightthickness=0, bd=0)
    warn_cv.pack(fill="x", pady=(16, 0))

    run_row = tk.Frame(col, bg=BG)
    run_row.pack(fill="x", pady=(8, 0))
    run_lbl = tk.Label(run_row, text="", font=f_run, fg=INK_2, bg=BG, anchor="w")
    run_lbl.pack(side="left", fill="x", expand=True)
    open_btn = tk.Label(run_row, text="打开文件夹", font=(ui, 11), fg=GREEN, bg=BG, cursor="hand2")

    log_host = tk.Frame(col, bg=BG)
    log_host.pack(fill="both", expand=True, pady=(8, 0))
    log_cv = tk.Canvas(log_host, bg=BG, highlightthickness=0, bd=0)
    log_cv.place(relx=0, rely=0, relwidth=1, relheight=1)
    inner = tk.Frame(log_host, bg=LOG_BG)
    inner.place(x=14, y=12, relwidth=1, relheight=1, width=-28, height=-24)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(
        "Dark.Vertical.TScrollbar",
        background="#3D4C63", troughcolor=LOG_BG, bordercolor=LOG_BG, arrowcolor="#FFFFFF",
        gripcount=0, borderwidth=0, lightcolor="#3D4C63", darkcolor="#3D4C63",
    )
    style.map("Dark.Vertical.TScrollbar", background=[("active", "#536480")])
    scroll = ttk.Scrollbar(inner, orient="vertical", style="Dark.Vertical.TScrollbar")
    txt = tk.Text(
        inner, wrap="word", font=(mono, 16), fg=LOG_FG, bg=LOG_BG, relief="flat", bd=0,
        padx=12, pady=10, spacing1=2, spacing3=4, height=8, state="disabled",
        highlightthickness=0, cursor="arrow", insertbackground=LOG_FG,
        selectbackground="#245078", selectforeground="#FFFFFF",
        yscrollcommand=scroll.set,
    )
    scroll.configure(command=txt.yview)
    scroll.pack(side="right", fill="y")
    txt.pack(side="left", fill="both", expand=True)
    txt.tag_configure("bad", foreground=LOG_BAD)
    txt.tag_configure("ok", foreground=LOG_OK)
    txt.tag_configure("hint", foreground=LOG_HINT)
    txt.tag_configure("result", foreground=LOG_OK, font=f_result)
    txt.tag_configure("stat", foreground=LOG_OK, font=f_stat)
    txt.tag_configure("guide_h", foreground="#FFFFFF", font=f_guide_h)

    def sync_log(e=None):
        w = log_host.winfo_width()
        h = log_host.winfo_height()
        if w < 20 or h < 20:
            return
        log_cv.delete("bg")
        _round_rect(log_cv, 1, 1, w - 1, h - 1, 14, fill=LOG_BG, outline="", tags="bg")

    log_host.bind("<Configure>", sync_log)

    def on_wheel(e):
        txt.yview_scroll(int(-e.delta / 120), "units")

    txt.bind("<MouseWheel>", on_wheel)

    IDLE_HINT = (
        "选好 ECN 和原理图 BOM 后，点「生成 ECO 脚本」。\n"
        "通过校验才会写出文件；未通过会列出原因。\n"
        "\n"
        "导入前请备份原理图，断开 Layout，并关闭在线 DRC。"
    )
    DONE_STEPS = (
        "1. 备份原理图\n"
        "2. 断开 Layout\n"
        "3. 关闭在线 DRC\n"
        "4. 文件 → 导入 → 文件类型 .ECO → 选择此次生成的 ECO 文件\n"
        "\n"
    )
    DONE_REV = (
        "原理图改完后，重新导出 BOM，点第 5 步选这份新 BOM。\n"
        "先关掉已打开的 ECN 工作簿（需要本机 Excel）。\n"
        "变更表没有数据行即通过。点「打开文件」核对照备份。"
    )

    lines: list[str] = []

    def log(s: str = ""):
        lines.append(s)
        tag = "ok" if s.startswith("PASS") else "bad" if s.startswith(("FAIL", "停止", "错误", "已改名", "  - ")) else ""
        txt.configure(state="normal")
        txt.insert("end", s + "\n", tag)
        txt.configure(state="disabled")
        txt.see("end")
        root.update_idletasks()

    def clear_log(hint: bool = False):
        lines.clear()
        txt.configure(state="normal", wrap="word")
        txt.delete("1.0", "end")
        if hint:
            txt.insert("end", IDLE_HINT, "hint")
        txt.configure(state="disabled")
        open_btn.pack_forget()

    def show_done_guide(eco_name: str, brief: str = ""):
        clear_log()
        txt.configure(state="normal", wrap="word")
        txt.insert("end", f"已生成  {eco_name}\n", "result")
        if brief and brief != "未生成":
            txt.insert("end", f"{brief}\n", "stat")
        txt.insert("end", "对照表在同目录 .map.csv。\n\n")
        txt.insert("end", "导入到 PADS Logic\n", "guide_h")
        txt.insert("end", DONE_STEPS)
        txt.insert("end", "反向检查\n", "guide_h")
        txt.insert("end", DONE_REV)
        txt.configure(state="disabled")
        txt.see("1.0")

    def set_verdict(ok: bool | None, text: str):
        if ok is None:
            run_lbl.configure(text=text, fg=INK_2, font=f_run)
        elif ok:
            run_lbl.configure(text="PASS  " + text, fg=GREEN, font=f_run_ok)
        else:
            run_lbl.configure(text="FAIL  " + text, fg=FAIL_C, font=f_run_ok)

    def summarize() -> str:
        joined = "\n".join(lines)
        m = re.search(r"modify (\d+) modify-via-NC (\d+) delete (\d+) add (\d+)", joined)
        c = re.search(r"修改 (\d+)，删除 (\d+)，新增 (\d+)", joined)
        k = re.search(r"(?:^skip\s+|跳过 )(\d+)", joined, re.M)
        p = re.search(r"(?:^pkg-change\s+|封装变化 )(\d+)", joined, re.M)
        if c:
            parts = [f"修改 {c.group(1)}", f"删除 {c.group(2)}", f"新增 {c.group(3)}"]
        elif m:
            mod = int(m.group(1)) + int(m.group(2))
            parts = [f"修改 {mod}", f"删除 {m.group(3)}", f"新增 {m.group(4)}"]
        else:
            return "未生成"
        if k and int(k.group(1)):
            parts.append(f"跳过 {k.group(1)}")
        if p and int(p.group(1)):
            parts.append(f"封装变化 {p.group(1)}")
        return " · ".join(parts)

    lib_cache: dict = {"key": None, "view": None}

    def lib_view() -> dict:
        lib = state["lib"]
        blank = {"n_both": None, "n_name_only": None}
        if not lib.exists():
            return {"ok": False, "brief": "未找到器件库", "detail": "", "reason": "找不到器件库", **blank}
        try:
            st = lib.stat()
            key = (str(lib.resolve()), st.st_mtime_ns, st.st_size)
        except OSError as e:
            return {"ok": False, "brief": "器件库无法读取", "detail": str(e), "reason": str(e), **blank}
        if lib_cache["key"] == key:
            return lib_cache["view"]
        try:
            audit, _pn, _names = ecn_to_pads_eco.read_library(lib)
        except Exception as e:
            plain = user_file_error(e)
            detail = plain or f"错误: {type(e).__name__}: {e}"
            view = {"ok": False, "brief": "器件库无法读取", "detail": detail, "reason": plain or str(e), **blank}
        else:
            reason = ecn_to_pads_eco.library_block_reason(audit)
            view = {
                "ok": reason is None,
                "brief": ecn_to_pads_eco.library_audit_brief(audit),
                "detail": ecn_to_pads_eco.format_library_audit(audit),
                "reason": reason or "",
                "n_both": audit.n_both,
                "n_name_only": audit.n_name_only,
            }
        lib_cache["key"] = key
        lib_cache["view"] = view
        return view

    def show_lib_gate():
        view = lib_view()
        if view["detail"]:
            for line in view["detail"].splitlines():
                log(line)
        if view["ok"]:
            set_verdict(None, "库可生成 · " + view["brief"])
        else:
            if view["reason"]:
                log("停止: " + view["reason"])
            set_verdict(False, "器件库信息不齐全，不能生成")

    def phases():
        flags = [
            bool(state.get("lib_step_done")),
            bool(state["ecn"]),
            bool(state.get("bom")),
            bool(state.get("gen_ok")),
            bool(state.get("rev_ok")),
        ]
        current = next((i for i, done in enumerate(flags) if not done), None)
        out = []
        for i, done in enumerate(flags):
            if done:
                out.append("done")
            elif i == current:
                out.append("current")
            else:
                out.append("wait")
        return out

    def badge_for(name: str) -> str:
        ext = Path(name).suffix.lower()
        if ext == ".xlsm":
            return "XLSM"
        if ext == ".csv":
            return "CSV"
        return "XLS"

    def hit_at(x, y):
        for x1, y1, x2, y2, tag in state["hits"]:
            if x1 <= x <= x2 and y1 <= y <= y2:
                return tag
        return None

    def paint_stepper(_e=None):
        if state["painting"]:
            return
        w = stepper.winfo_width()
        if w < 80:
            return
        state["painting"] = True
        try:
            stepper.delete("all")
            phase = phases()
            enabled = state["ready"] and not state["busy"]
            hover = state["hover"]
            specs = [
                {"kind": "file", "key": "lib", "phase": 0, "title": "更新器件库",
                 "filled": state["lib"].exists(), "filename": state["lib"].name if state["lib"].exists() else "",
                 "bad": state["lib"].exists() and not state["lib_ok"]},
                {"kind": "file", "key": "ecn", "phase": 1, "title": "导入ECN",
                 "filled": bool(state["ecn"]), "filename": state["ecn"].name if state["ecn"] else "", "bad": False},
                {"kind": "file", "key": "bom", "phase": 2, "title": "导入原理图BOM",
                 "filled": bool(state["bom"]), "filename": state["bom"].name if state["bom"] else "",
                 "bad": False},
                {"kind": "gen", "phase": 3},
                {"kind": "file", "key": "rev", "phase": 4, "title": "反向检查",
                 "filled": bool(state["rev_name"]), "filename": state["rev_name"] or "", "bad": False},
            ]
            gap = 16
            y = 2
            geom = []
            for row in specs:
                h = 52 if row["kind"] == "gen" else 84
                geom.append((y, h, row))
                y += h + gap
            total_h = max(1, y - gap + 4)
            if int(stepper.cget("height")) != total_h:
                stepper.configure(height=total_h)
            centers = [(y0 + h / 2, row["phase"]) for y0, h, row in geom]

            def rail_gap(name):
                return 16 if name == "current" else 13

            for (y1, p1), (y2, p2) in zip(centers, centers[1:]):
                stepper.create_line(
                    16, y1 + rail_gap(phase[p1]), 16, y2 - rail_gap(phase[p2]),
                    fill=GREEN if phase[p1] == "done" else "#C5CCD6", width=1,
                )
            hits = []

            def dot_photo(kind):
                if kind not in state["dot_img"]:
                    from PIL import ImageTk
                    state["dot_img"][kind] = ImageTk.PhotoImage(_dot_rgba(kind))
                return state["dot_img"][kind]

            for y0, h, row in geom:
                cy = y0 + h / 2
                kind = phase[row["phase"]]
                # 蓝描边只给下一张要填的文件卡。未就绪的「生成」仍是灰按钮。
                if row["kind"] == "gen" and kind == "current" and not enabled:
                    kind = "wait"
                dot_kind = "gen_on" if row["kind"] == "gen" and kind == "current" else kind
                stepper.create_image(16, cy, image=dot_photo(dot_kind))
                if dot_kind == "current":
                    stepper.create_text(16, cy, text=str(row["phase"] + 1), fill=CUR_DOT, font=f_num)
                elif dot_kind == "gen_on":
                    stepper.create_text(16, cy, text=str(row["phase"] + 1), fill=GREEN, font=f_num)
                elif dot_kind == "wait":
                    stepper.create_text(16, cy, text=str(row["phase"] + 1), fill=INK_3, font=f_num)
                x1, x2 = 40, w - 6
                if row["kind"] == "gen":
                    fill = GREEN_DK if enabled and state["pressed"] else (GREEN if enabled else "#E6E8EC")
                    fg = WHITE if enabled else INK_3
                    if enabled:
                        _round_rect(stepper, x1, y0 + 3, x2, y0 + h + 1, 10, fill="#E3EBE8", outline="")
                    _round_rect(stepper, x1, y0, x2, y0 + h - 2, 10, fill=fill, outline="")
                    stepper.create_text((x1 + x2) / 2, y0 + (h - 2) / 2, text="生成 ECO 脚本", fill=fg, font=f_gen)
                    hits.append((x1, y0, x2, y0 + h, "gen"))
                    continue
                filled = row["filled"]
                clickable = row.get("clickable", True)
                hot = clickable and hover == row["key"]
                is_current = kind == "current"
                if is_current:
                    _round_rect(stepper, x1 - 2, y0 - 2, x2 + 2, y0 + h + 2, 14, fill=CUR_RING, outline="")
                    _round_rect(stepper, x1, y0, x2, y0 + h, 12, fill=WHITE, outline=CUR_LINE, width=1.5)
                elif filled:
                    _round_rect(stepper, x1, y0, x2, y0 + h, 12, fill=GREEN_SOFT, outline=GREEN if hot else GREEN_LINE, width=1)
                else:
                    _round_rect(stepper, x1, y0, x2, y0 + h, 12, fill=WHITE, outline="")
                    _stroke_round(stepper, x1, y0, x2, y0 + h, 12, GREEN if hot else DASH, 1, (3, 3))
                bh = 34
                right = x2 - 16
                button_specs = []
                if clickable:
                    if row["key"] == "rev" and state.get("rev_book") and Path(state["rev_book"]).is_file():
                        button_specs.append((max(72, f_ghost.measure("打开文件") + 28), "打开文件", "rev_open"))
                    button_text = row.get("button") or ("更换" if filled else "选择")
                    button_specs.append((72, button_text, None))
                left_edge = right
                for width, text, tag in reversed(button_specs):
                    bx2 = left_edge
                    bx1 = bx2 - width
                    by1 = y0 + (h - bh) / 2
                    _round_rect(
                        stepper, bx1, by1, bx2, by1 + bh, 8,
                        fill="#F3F4F6" if tag and hover == tag else WHITE,
                        outline=LINE, width=1,
                    )
                    stepper.create_text(
                        (bx1 + bx2) / 2, by1 + bh / 2, text=text,
                        fill=GREEN if tag == "rev_open" else INK, font=f_ghost,
                    )
                    if tag:
                        hits.insert(0, (bx1, by1, bx2, by1 + bh, tag))
                    left_edge = bx1 - 8
                file_right = left_edge + 8
                lead = max(52, f_badge.measure("XLSM") + 18)
                text_l = x1 + 16
                if filled or (row.get("badge") and row.get("filename")):
                    label = row.get("badge") or badge_for(row["filename"])
                    tw = max(44, f_badge.measure(label) + 18)
                    badge_h = 26
                    ty1 = y0 + (h - badge_h) / 2
                    bx = text_l + (lead - tw) / 2
                    _round_rect(stepper, bx, ty1, bx + tw, ty1 + badge_h, 6, fill="#D8F3E8", outline="")
                    stepper.create_text(bx + tw / 2, ty1 + badge_h / 2, text=label, fill=GREEN, font=f_badge)
                elif clickable:
                    px = text_l + lead / 2
                    stepper.create_line(px - 6, cy, px + 6, cy, fill=INK_3, width=1)
                    stepper.create_line(px, cy - 6, px, cy + 6, fill=INK_3, width=1)
                text_l += lead + 14
                max_w = max(20, file_right - 12 - text_l)
                title = _ellipsize(f_card, row["title"], max_w)
                if filled:
                    stepper.create_text(text_l, cy - 14, text=title, anchor="w", fill=INK, font=f_card)
                    name_fg = FAIL_C if row["bad"] else INK_2
                    stepper.create_text(text_l, cy + 16, text=_ellipsize(f_file, row["filename"], max_w), anchor="w", fill=name_fg, font=f_file)
                elif row.get("hint"):
                    stepper.create_text(text_l, cy - 14, text=title, anchor="w", fill=INK, font=f_card)
                    stepper.create_text(text_l, cy + 16, text=_ellipsize(f_file, row["hint"], max_w), anchor="w", fill=INK_2, font=f_file)
                else:
                    stepper.create_text(text_l, cy, text=title, anchor="w", fill=INK, font=f_card)
                if clickable:
                    hits.append((x1, y0, x2, y0 + h, row["key"]))
            state["hits"] = hits
        finally:
            state["painting"] = False

    def paint_warn(_e=None):
        w = warn_cv.winfo_width()
        if w < 80:
            return
        view = state["view"] or {}
        text = "导入前请备份 .sch，断开 Layout，并关闭 Online DRC。"
        n = view.get("n_name_only") or 0
        if n:
            text += f"缺料号 {n} 项建议先核对器件库。"
        lines_w = _wrap_text(f_warn, text, w - 28)
        h = 16 + 18 * len(lines_w) + 6
        if int(warn_cv.cget("height")) != h:
            warn_cv.configure(height=h)
        warn_cv.delete("all")
        _round_rect(warn_cv, 1, 1, w - 2, h - 2, 10, fill=AMBER_BG, outline=AMBER_LINE, width=1)
        y = 10
        for line in lines_w:
            warn_cv.create_text(14, y, text=line, anchor="nw", fill=AMBER, font=f_warn)
            y += 18

    def refresh():
        view = lib_view()
        state["lib_ok"] = view["ok"]
        state["view"] = view
        state["ready"] = bool(
            state["ecn"] and state["bom"] and state["lib"].exists() and state["lib_ok"]
        )
        paint_stepper()
        paint_warn()

    def on_gen():
        clear_log()
        set_verdict(None, "正在生成并校验…")
        state["busy"] = True
        paint_stepper()
        root.update_idletasks()
        ok = False
        try:
            ok = do_gen(
                state["ecn"], state["bom"], state["lib"], None, log,
                lambda rows: ask_continue_skips(root, rows),
                confirm_pkg=lambda rows: ask_continue_pkg(root, rows, state["ecn"]),
            )
        except Exception:
            ok = False
            log("错误: 程序异常")
            log(traceback.format_exc())
        state["busy"] = False
        if ok:
            state["gen_ok"] = True
            state["rev_ok"] = False
            state["last_out"] = state["ecn"].with_suffix(".eco")
            had_bad_skip = any(s.startswith("异常跳过 ") for s in lines)
            brief = summarize()
            show_done_guide(state["last_out"].name, brief)
            if had_bad_skip:
                log("")
                log("生成时有异常跳过。被跳过的位号和原因在同目录 .map.csv，请先核对再导入。")
            set_verdict(True, brief + f"  →  {state['last_out'].name}")
            open_btn.pack(side="right")
        else:
            state["gen_ok"] = False
            set_verdict(False, "未写出可导入的 ECO，见下方原因")
        refresh()

    def on_swap(src: Path | None = None):
        if src is None:
            p = filedialog.askopenfilename(title="选择器件库 CSV", filetypes=[("CSV", "*.csv"), ("所有文件", "*.*")])
            if not p:
                return
            src = Path(p)
        elif src.suffix.lower() != ".csv":
            set_verdict(False, "更新器件库请选择 CSV")
            return
        dst = exe_dir() / "library" / LIB_NAME
        clear_log()
        try:
            n = ecn_to_pads_eco.import_library_csv(src, dst)
        except Exception as e:
            log(f"更新器件库失败，原 xlsx 未改: {e}")
            set_verdict(False, "器件库未更新")
            refresh()
            return
        state["lib"] = dst
        state["gen_ok"] = False
        state["rev_ok"] = False
        state["rev_name"] = None
        state["rev_book"] = None
        show_lib_gate()
        state["lib_step_done"] = bool(lib_view().get("ok"))
        run_lbl.configure(text=f"已写入 {dst.name} ← {src.name} · {n} 行", fg=INK, font=f_run)
        refresh()

    def choose_file(key, title):
        def action(dropped: Path | None = None):
            path = dropped
            if path is None:
                p = filedialog.askopenfilename(title=title, filetypes=[("Excel", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
                if not p:
                    return
                path = Path(p)
            elif path.suffix.lower() not in {".xlsx", ".xlsm"}:
                set_verdict(False, "请选择 Excel 文件（.xlsx / .xlsm）")
                return
            state[key] = path
            state["gen_ok"] = False
            state["rev_ok"] = False
            state["rev_name"] = None
            state["rev_book"] = None
            refresh()
        return action

    pick_ecn = choose_file("ecn", "选择 ECN")
    pick_bom = choose_file("bom", "选择原理图 BOM")

    def on_reverse(src: Path | None = None):
        if src is None:
            p = filedialog.askopenfilename(
                title="选择导入 ECO 后重新导出的原理图 BOM",
                filetypes=[("Excel", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
            )
            if not p:
                return
            src = Path(p)
        elif src.suffix.lower() not in {".xlsx", ".xlsm"}:
            set_verdict(False, "反向检查请选择 Excel BOM")
            return
        tool = state["ecn"]
        if not (tool and tool.suffix.lower() == ".xlsm" and tool.exists()):
            picked = filedialog.askopenfilename(
                title="选择 ECN 工具（Bom&ECN-changelist.xlsm）",
                filetypes=[("Excel 宏工作簿", "*.xlsm"), ("所有文件", "*.*")],
            )
            if not picked:
                return
            tool = Path(picked)
        clear_log()
        state["rev_book"] = None
        state["rev_name"] = src.name
        paint_stepper()
        set_verdict(None, "正在反向检查…")
        report = None
        try:
            report = run_reverse(src, tool, log)
        except Exception as e:
            log(f"反向检查失败: {e}")
            set_verdict(False, "反向检查未完成")
        if report is not None:
            if report.check_note:
                log(f"位号检查  {report.check_note}")
            for text in report.dialogs:
                log(f"宏: {text}")
            used = report.workbook.name if report.workbook else tool.name
            state["rev_name"] = used
            state["rev_ok"] = bool(report.passed)
            if report.passed:
                log(f"PASS  {report.summary}  ({used})")
            else:
                log(f"FAIL  {report.summary}  ({used})")
                for ref in report.refs[:40]:
                    log(f"  - {ref}")
                extra = len(report.refs) - 40
                if extra > 0:
                    log(f"  - …还有 {extra} 行")
                if not report.finished:
                    log("ECN 表没有重写，不能据此说 ECN 为空。")
            set_verdict(True if report.passed else False, report.summary)
        show_rev_file(report.workbook if report is not None else _rev_backup(tool))
        refresh()

    def _rev_backup(tool: Path) -> Path | None:
        """这次检查已经动过的备份。复制都没成功时不把旧文件当成这次的结果。"""
        touched = any(
            s.startswith("已复制 ") or s.startswith("使用已有备份") or "EBOM 已写入" in s
            for s in lines
        )
        if not touched:
            return None
        tool = Path(tool).resolve()
        if tool.stem.endswith(".反向检查"):
            return tool
        return tool.with_name(f"{tool.stem}.反向检查.xlsm")

    def show_rev_file(book: Path | None):
        if book is not None and Path(book).is_file():
            state["rev_book"] = Path(book)
            log(f"点第 5 步的「打开文件」打开 {state['rev_book'].name}")
        else:
            state["rev_book"] = None

    def on_open(_e=None):
        if state["last_out"]:
            os.startfile(str(state["last_out"].parent))

    def on_open_rev(_e=None):
        book = state.get("rev_book")
        if not book or not Path(book).is_file():
            log("找不到反向检查文件。")
            state["rev_book"] = None
            refresh()
            return
        try:
            os.startfile(str(book))
        except OSError as e:
            log(f"打不开 {Path(book).name}: {e}")

    def dispatch(tag):
        if tag == "lib":
            on_swap()
        elif tag == "ecn":
            pick_ecn()
        elif tag == "bom":
            pick_bom()
        elif tag == "rev_open":
            on_open_rev()
        elif tag == "rev":
            on_reverse()
        elif tag == "gen" and state["ready"] and not state["busy"]:
            on_gen()

    def on_press(e):
        tag = hit_at(e.x, e.y)
        state["press"] = tag
        if tag == "gen" and state["ready"] and not state["busy"]:
            state["pressed"] = True
            paint_stepper()

    def on_release(e):
        tag = hit_at(e.x, e.y)
        was = state["pressed"]
        state["pressed"] = False
        same = state["press"] == tag and tag
        state["press"] = None
        if same:
            dispatch(tag)
        elif was:
            paint_stepper()

    def on_motion(e):
        tag = hit_at(e.x, e.y)
        clickable = bool(tag) and not (tag == "gen" and not (state["ready"] and not state["busy"]))
        stepper.configure(cursor="hand2" if clickable else "")
        if tag != state["hover"]:
            state["hover"] = tag
            paint_stepper()

    def on_leave(_e):
        stepper.configure(cursor="")
        if state["hover"]:
            state["hover"] = None
            paint_stepper()

    stepper.bind("<ButtonPress-1>", on_press)
    stepper.bind("<ButtonRelease-1>", on_release)
    stepper.bind("<Motion>", on_motion)
    stepper.bind("<Leave>", on_leave)
    stepper.bind("<Configure>", paint_stepper)
    warn_cv.bind("<Configure>", paint_warn)
    open_btn.bind("<Button-1>", on_open)

    def on_drop(files, x, y):
        if not files:
            return
        sx = root.winfo_rootx() + int(x)
        sy = root.winfo_rooty() + int(y)
        tag = hit_at(sx - stepper.winfo_rootx(), sy - stepper.winfo_rooty())
        path = Path(files[0])
        if tag == "lib":
            on_swap(path)
        elif tag == "ecn":
            pick_ecn(path)
        elif tag == "bom":
            pick_bom(path)
        elif tag == "rev":
            on_reverse(path)

    def install_drop():
        if sys.platform != "win32":
            return
        import ctypes
        from ctypes import wintypes
        root.update_idletasks()
        hwnd = wintypes.HWND(root.winfo_id())
        user32 = ctypes.windll.user32
        shell32 = ctypes.windll.shell32
        user32.GetWindowLongPtrW.restype = ctypes.c_void_p
        user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.SetWindowLongPtrW.restype = ctypes.c_void_p
        user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        user32.CallWindowProcW.restype = ctypes.c_ssize_t
        user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM]
        shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
        shell32.DragQueryFileW.restype = wintypes.UINT
        shell32.DragFinish.argtypes = [wintypes.HANDLE]

        class POINT(ctypes.Structure):
            _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

        shell32.DragQueryPoint.argtypes = [wintypes.HANDLE, ctypes.POINTER(POINT)]
        shell32.DragQueryPoint.restype = wintypes.BOOL
        old = user32.GetWindowLongPtrW(hwnd, -4)
        if not old:
            return
        WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM)

        def proc(h, msg, wp, lp):
            if msg == 0x0233:
                try:
                    count = shell32.DragQueryFileW(wp, 0xFFFFFFFF, None, 0)
                    files = []
                    for i in range(count):
                        n = shell32.DragQueryFileW(wp, i, None, 0)
                        buf = ctypes.create_unicode_buffer(n + 1)
                        shell32.DragQueryFileW(wp, i, buf, n + 1)
                        files.append(buf.value)
                    pt = POINT()
                    shell32.DragQueryPoint(wp, ctypes.byref(pt))
                    shell32.DragFinish(wp)
                    root.after(0, lambda f=list(files), px=pt.x, py=pt.y: on_drop(f, px, py))
                except Exception:
                    try:
                        shell32.DragFinish(wp)
                    except Exception:
                        pass
                return 0
            return user32.CallWindowProcW(old, h, msg, wp, lp)

        cb = WNDPROC(proc)
        user32.SetWindowLongPtrW(hwnd, -4, ctypes.cast(cb, ctypes.c_void_p).value)
        shell32.DragAcceptFiles(hwnd, True)
        root._drop_cb = cb

    def install_chrome():
        """Replace the native frame with the in-window title bar."""
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32

        class RECT(ctypes.Structure):
            _fields_ = [("l", wintypes.LONG), ("t", wintypes.LONG), ("r", wintypes.LONG), ("b", wintypes.LONG)]

        def outer_hwnd():
            return user32.GetParent(root.winfo_id())

        def fix_taskbar():
            GWL_EXSTYLE = -20
            WS_EX_APPWINDOW = 0x00040000
            WS_EX_TOOLWINDOW = 0x00000080
            hwnd = outer_hwnd()
            user32.GetWindowLongPtrW.restype = ctypes.c_void_p
            user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
            user32.SetWindowLongPtrW.restype = ctypes.c_void_p
            user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
            style = int(user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE) or 0)
            style = (style & ~WS_EX_TOOLWINDOW) | WS_EX_APPWINDOW
            user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, style)
            user32.SetWindowPos(hwnd, None, 0, 0, 0, 0, 0x0027)

        def work_area():
            rect = RECT()
            user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0)
            return rect.l, rect.t, rect.r - rect.l, rect.b - rect.t

        def refresh_max():
            paint_caption(caption_paint["max"], "max", WHITE, TITLE_FG)

        def toggle_zoom(_e=None):
            if chrome["zoom"]:
                root.geometry(chrome["geom"] or "1000x890+80+30")
                chrome["zoom"] = False
            else:
                chrome["geom"] = root.geometry()
                x, y, w, h = work_area()
                root.geometry(f"{w}x{h}+{x}+{y}")
                chrome["zoom"] = True
            refresh_max()
            return "break"

        def do_min(_e=None):
            chrome["geom_min"] = root.geometry()
            root.overrideredirect(False)
            root.iconify()
            return "break"

        def on_map(e):
            if e.widget is not root or chrome.get("mapping"):
                return
            if str(root.state()) == "iconic":
                return
            chrome["mapping"] = True
            try:
                root.overrideredirect(True)
                fix_taskbar()
            finally:
                chrome["mapping"] = False

        def start_drag(e):
            if chrome["zoom"]:
                chrome["drag"] = "zoom"
                chrome["drag_ratio"] = e.x_root - root.winfo_rootx()
                return
            chrome["drag"] = (e.x_root - root.winfo_x(), e.y_root - root.winfo_y())

        def on_drag(e):
            drag = chrome.get("drag")
            if drag == "zoom":
                ratio = chrome.get("drag_ratio") or 0
                toggle_zoom()
                w = root.winfo_width()
                x = e.x_root - min(max(ratio, 40), w - 80)
                y = max(0, e.y_root - 14)
                root.geometry(f"+{x}+{y}")
                chrome["drag"] = (e.x_root - root.winfo_x(), e.y_root - root.winfo_y())
                return
            if not isinstance(drag, tuple):
                return
            dx, dy = drag
            root.geometry(f"+{e.x_root - dx}+{e.y_root - dy}")

        def end_drag(_e):
            if chrome.get("drag") != "zoom":
                chrome["drag"] = None

        for widget in (bar, title_lbl):
            widget.bind("<ButtonPress-1>", start_drag)
            widget.bind("<B1-Motion>", on_drag)
            widget.bind("<ButtonRelease-1>", end_drag)
            widget.bind("<Double-Button-1>", toggle_zoom)
        btn_min.bind("<ButtonRelease-1>", do_min)
        btn_max.bind("<ButtonRelease-1>", toggle_zoom)
        btn_close.bind("<ButtonRelease-1>", lambda _e: root.destroy())
        root.bind("<Map>", on_map)
        root.bind("<Alt-F4>", lambda _e: root.destroy())
        root.protocol("WM_DELETE_WINDOW", root.destroy)

        edge = {"name": None, "x": 0, "y": 0, "gx": 0, "gy": 0, "w": 0, "h": 0}
        grips = {}

        def begin_resize(name):
            def start(e):
                if chrome["zoom"]:
                    return "break"
                edge.update(name=name, x=e.x_root, y=e.y_root, gx=root.winfo_x(), gy=root.winfo_y(),
                            w=root.winfo_width(), h=root.winfo_height())
                return "break"

            def drag(e):
                if not edge["name"] or chrome["zoom"]:
                    return "break"
                dx, dy = e.x_root - edge["x"], e.y_root - edge["y"]
                x, y, w, h = edge["gx"], edge["gy"], edge["w"], edge["h"]
                if "e" in name:
                    w = max(880, w + dx)
                if "s" in name:
                    h = max(750, h + dy)
                if "w" in name:
                    nw = max(880, w - dx)
                    x += w - nw
                    w = nw
                if "n" in name:
                    nh = max(750, h - dy)
                    y += h - nh
                    h = nh
                root.geometry(f"{int(w)}x{int(h)}+{int(x)}+{int(y)}")
                return "break"

            def stop(_e):
                edge["name"] = None

            return start, drag, stop

        cursors = {
            "n": "sb_v_double_arrow", "s": "sb_v_double_arrow",
            "e": "sb_h_double_arrow", "w": "sb_h_double_arrow",
            "nw": "size_nw_se", "se": "size_nw_se", "ne": "size_ne_sw", "sw": "size_ne_sw",
        }
        for name, cursor in cursors.items():
            color = WHITE if name in {"n", "nw", "ne"} else BG
            g = tk.Frame(root, bg=color, cursor=cursor, highlightthickness=0, bd=0)
            start, drag, stop = begin_resize(name)
            g.bind("<ButtonPress-1>", start)
            g.bind("<B1-Motion>", drag)
            g.bind("<ButtonRelease-1>", stop)
            grips[name] = g

        def layout_grips(_e=None):
            w, h = root.winfo_width(), root.winfo_height()
            if w < 40 or h < 40:
                return
            m = 6
            grips["n"].place(x=m, y=0, width=w - 2 * m, height=m)
            grips["s"].place(x=m, y=h - m, width=w - 2 * m, height=m)
            grips["w"].place(x=0, y=m, width=m, height=h - 2 * m)
            grips["e"].place(x=w - m, y=BAR_H, width=m, height=h - BAR_H - m)
            grips["nw"].place(x=0, y=0, width=m, height=m)
            grips["ne"].place(x=w - m, y=0, width=m, height=m)
            grips["sw"].place(x=0, y=h - m, width=m, height=m)
            grips["se"].place(x=w - m, y=h - m, width=m, height=m)
            for g in grips.values():
                g.lift()

        root.bind("<Configure>", layout_grips, add="+")
        root.overrideredirect(True)
        root.update_idletasks()
        fix_taskbar()
        root.withdraw()
        root.after(10, root.deiconify)

    try:
        install_chrome()
    except Exception:
        pass

    try:
        install_drop()
    except Exception:
        pass

    refresh()
    clear_log(hint=True)
    if state["lib"].exists() and not state["lib_ok"]:
        clear_log()
        show_lib_gate()
    root.mainloop()
    return 0


def main() -> int:
    if len(sys.argv) > 1:
        if sys.stdout is None:  # windowed exe called with args: no console to print to
            import os
            sys.stdout = sys.stderr = open(os.devnull, "w", encoding="utf-8")
        return cli(sys.argv[1:])
    return gui()


if __name__ == "__main__":
    sys.exit(main())
