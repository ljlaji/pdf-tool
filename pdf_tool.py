"""PDF 工具：查看 / 拆分 / 合并（PDF+图片） / 旋转 / 保存

依赖：pip install pypdf pypdfium2 Pillow tkinterdnd2
运行：python pdf_tool.py
"""

from __future__ import annotations

import io
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pypdfium2 as pdfium
from PIL import Image, ImageTk
from pypdf import PdfReader, PdfWriter

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    DND_FILES = None
    TkinterDnD = None


IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".gif", ".webp"}


def image_to_pdf_bytes(path):
    """把图片转换为单页 PDF，返回字节。"""
    with Image.open(path) as im:
        if im.mode in ("RGBA", "LA", "P"):
            im = im.convert("RGB")
        elif im.mode != "RGB":
            im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, format="PDF", resolution=150.0)
        return buf.getvalue()


def register_drop(widget, callback):
    """把 widget 注册为文件拖拽目标；callback 接收路径列表。"""
    if not DND_AVAILABLE:
        return
    try:
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", lambda e: callback(list(widget.tk.splitlist(e.data))))
    except Exception:
        pass


# ---------- 工具函数 ----------

def parse_page_ranges(expr, total):
    """解析页码表达式，返回每个输出文件对应的页码列表（0-based）。"""
    expr = expr.strip()
    if not expr:
        return [[i] for i in range(total)]

    groups = []
    for part in expr.split(";"):
        part = part.strip()
        if not part:
            continue
        pages = []
        for token in part.split(","):
            token = token.strip()
            if not token:
                continue
            m = re.fullmatch(r"(\d+)\s*-\s*(\d+)", token)
            if m:
                a, b = int(m.group(1)), int(m.group(2))
                if a > b:
                    a, b = b, a
                pages.extend(range(a - 1, b))
            elif token.isdigit():
                pages.append(int(token) - 1)
            else:
                raise ValueError(f"无法解析的页码片段：{token}")
        for p in pages:
            if p < 0 or p >= total:
                raise ValueError(f"页码 {p + 1} 超出范围（共 {total} 页）")
        if pages:
            groups.append(pages)
    return groups


# ---------- 查看 / 编辑 标签页 ----------

THUMB_WIDTH = 140
ZOOM_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 1.0, 1.25, 1.5, 1.75, 2.0, 2.5, 3.0, 4.0]


class ViewerTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.doc = None              # pypdfium2.PdfDocument
        self.path = None             # 当前文件路径
        self.page_count = 0
        self.current = 0             # 当前页 0-based
        self.zoom_idx = 5            # 初始 1.0
        self.fit_width = False
        self.rotations = {}          # {page_idx: 0/90/180/270}
        self.thumb_refs = []         # 防止 GC
        self.thumb_items = []        # canvas 上的 item id 列表
        self.thumb_highlight = None
        self.page_img_ref = None
        self._resize_after = None

        self._build_toolbar()
        self._build_body()
        self._build_statusbar()

        self._update_ui_state()

        # 拖拽打开
        register_drop(self, self._on_drop)
        register_drop(self.page_canvas, self._on_drop)
        register_drop(self.thumb_canvas, self._on_drop)

    def _on_drop(self, paths):
        for p in paths:
            if p.lower().endswith(".pdf") and os.path.isfile(p):
                self.load(p)
                return

    # ----- UI 构建 -----
    def _build_toolbar(self):
        tb = ttk.Frame(self, padding=(6, 4))
        tb.pack(fill="x")

        def add(label, cmd):
            b = ttk.Button(tb, text=label, command=cmd, width=8)
            b.pack(side="left", padx=2)
            return b

        self.btn_open = add("打开", self.on_open)
        self.btn_save = add("保存", self.on_save)
        self.btn_saveas = add("另存为", self.on_save_as)

        ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=6)

        self.btn_prev = add("◀ 上一页", self.prev_page)
        self.page_var = tk.StringVar(value="0 / 0")
        self.page_entry = ttk.Entry(tb, textvariable=self.page_var, width=10, justify="center")
        self.page_entry.pack(side="left", padx=2)
        self.page_entry.bind("<Return>", self.on_page_entry)
        self.btn_next = add("下一页 ▶", self.next_page)

        ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=6)

        self.btn_zoom_out = add("缩小 −", self.zoom_out)
        self.zoom_var = tk.StringVar(value="100%")
        ttk.Label(tb, textvariable=self.zoom_var, width=6, anchor="center").pack(side="left")
        self.btn_zoom_in = add("放大 ＋", self.zoom_in)
        self.btn_fit = add("适合宽度", self.toggle_fit)

        ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=6)

        self.btn_rot_l = add("↺ 左旋", lambda: self.rotate_current(-90))
        self.btn_rot_r = add("↻ 右旋", lambda: self.rotate_current(90))

    def _build_body(self):
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)

        # 左侧缩略图
        left = ttk.Frame(body, width=THUMB_WIDTH + 30)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        ttk.Label(left, text="页面", padding=(6, 4)).pack(anchor="w")

        thumb_wrap = ttk.Frame(left)
        thumb_wrap.pack(fill="both", expand=True)
        self.thumb_canvas = tk.Canvas(thumb_wrap, bg="#3a3a3a", highlightthickness=0, width=THUMB_WIDTH + 12)
        self.thumb_scroll = ttk.Scrollbar(thumb_wrap, orient="vertical", command=self.thumb_canvas.yview)
        self.thumb_canvas.configure(yscrollcommand=self.thumb_scroll.set)
        self.thumb_canvas.pack(side="left", fill="both", expand=True)
        self.thumb_scroll.pack(side="right", fill="y")
        self.thumb_canvas.bind("<Button-1>", self.on_thumb_click)
        self.thumb_canvas.bind("<MouseWheel>", lambda e: self.thumb_canvas.yview_scroll(int(-e.delta / 120), "units"))

        # 中央主显示
        center = ttk.Frame(body)
        center.pack(side="left", fill="both", expand=True)

        self.page_canvas = tk.Canvas(center, bg="#525659", highlightthickness=0)
        hbar = ttk.Scrollbar(center, orient="horizontal", command=self.page_canvas.xview)
        vbar = ttk.Scrollbar(center, orient="vertical", command=self.page_canvas.yview)
        self.page_canvas.configure(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        hbar.pack(side="bottom", fill="x")
        self.page_canvas.pack(side="left", fill="both", expand=True)
        self.page_canvas.bind("<Configure>", self._on_canvas_resize)
        self.page_canvas.bind("<MouseWheel>", self._on_wheel)
        self.page_canvas.bind("<Control-MouseWheel>", self._on_ctrl_wheel)

    def _build_statusbar(self):
        self.status = tk.StringVar(value="未打开文件")
        bar = ttk.Frame(self, padding=(8, 2))
        bar.pack(fill="x", side="bottom")
        ttk.Label(bar, textvariable=self.status, foreground="#555").pack(side="left")

    # ----- 状态 -----
    def _update_ui_state(self):
        has = self.doc is not None
        state = "normal" if has else "disabled"
        for b in (self.btn_save, self.btn_saveas, self.btn_prev, self.btn_next,
                  self.btn_zoom_in, self.btn_zoom_out, self.btn_fit,
                  self.btn_rot_l, self.btn_rot_r):
            b.config(state=state)

    # ----- 打开 / 保存 -----
    def on_open(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self.load(path)

    def load(self, path):
        try:
            doc = pdfium.PdfDocument(path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开：{e}")
            return
        if self.doc is not None:
            try:
                self.doc.close()
            except Exception:
                pass
        self.doc = doc
        self.path = path
        self.page_count = len(doc)
        self.current = 0
        self.rotations = {}
        self._update_ui_state()
        self._render_thumbnails()
        self._render_page()
        self._update_status()

    def on_save(self):
        if not self.doc:
            return
        if not self.rotations:
            messagebox.showinfo("提示", "没有需要保存的修改")
            return
        self._save_to(self.path, overwrite_prompt=False)

    def on_save_as(self):
        if not self.doc:
            return
        init = os.path.splitext(os.path.basename(self.path))[0] + "_edited.pdf"
        out = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=init,
            initialdir=os.path.dirname(self.path) if self.path else None,
        )
        if out:
            self._save_to(out, overwrite_prompt=False)

    def _save_to(self, out_path, overwrite_prompt):
        try:
            # 覆盖同一文件时，先写临时文件再替换
            writing_over_self = os.path.abspath(out_path) == os.path.abspath(self.path or "")
            tmp_path = out_path + ".tmp" if writing_over_self else out_path

            reader = PdfReader(self.path)
            writer = PdfWriter()
            for i, page in enumerate(reader.pages):
                delta = self.rotations.get(i, 0) % 360
                if delta:
                    page.rotate(delta)
                writer.add_page(page)
            with open(tmp_path, "wb") as f:
                writer.write(f)

            if writing_over_self:
                # 关闭 pdfium 句柄后再替换
                try:
                    self.doc.close()
                except Exception:
                    pass
                self.doc = None
                os.replace(tmp_path, out_path)
                # 重新打开
                self.doc = pdfium.PdfDocument(out_path)
                self.rotations = {}
                self._render_thumbnails()
                self._render_page()
            else:
                self.rotations = {}
                self._refresh_thumb_badges()

            messagebox.showinfo("完成", f"已保存：\n{out_path}")
            self._update_status()
        except Exception as e:
            messagebox.showerror("错误", str(e))

    # ----- 渲染 -----
    def _render_thumbnails(self):
        self.thumb_canvas.delete("all")
        self.thumb_refs = []
        self.thumb_items = []
        y = 6
        for i in range(self.page_count):
            img = self._render_page_image(i, scale_by_width=THUMB_WIDTH)
            tkimg = ImageTk.PhotoImage(img)
            self.thumb_refs.append(tkimg)
            x = 6
            rect = self.thumb_canvas.create_rectangle(
                x - 2, y - 2, x + img.width + 2, y + img.height + 2,
                outline="", width=2, tags=(f"thumb_rect_{i}",)
            )
            item = self.thumb_canvas.create_image(x, y, anchor="nw", image=tkimg, tags=(f"thumb_{i}",))
            label = self.thumb_canvas.create_text(
                x + img.width / 2, y + img.height + 10,
                text=str(i + 1), fill="#e0e0e0", font=("Segoe UI", 9),
                tags=(f"thumb_label_{i}",),
            )
            self.thumb_items.append((rect, item, label, img.width, img.height))
            y += img.height + 22
        self.thumb_canvas.configure(scrollregion=(0, 0, THUMB_WIDTH + 12, y))
        self._highlight_current_thumb()

    def _refresh_thumb_badges(self):
        self._highlight_current_thumb()

    def _highlight_current_thumb(self):
        for i, (rect, *_ ) in enumerate(self.thumb_items):
            color = "#4a90e2" if i == self.current else ""
            self.thumb_canvas.itemconfig(rect, outline=color)
        # 滚动到可见
        if 0 <= self.current < len(self.thumb_items):
            rect = self.thumb_items[self.current][0]
            x1, y1, x2, y2 = self.thumb_canvas.coords(rect)
            total = self.thumb_canvas.bbox("all")
            if total:
                top = y1 / total[3]
                self.thumb_canvas.yview_moveto(max(0, top - 0.1))

    def _render_page_image(self, idx, scale=None, scale_by_width=None):
        page = self.doc[idx]
        delta = self.rotations.get(idx, 0) % 360
        if scale_by_width:
            w, h = page.get_size()
            # 页面当前渲染尺寸考虑旋转
            if delta in (90, 270):
                w, h = h, w
            scale = scale_by_width / w
        if scale is None:
            scale = 1.0
        pil = page.render(scale=scale, rotation=delta).to_pil()
        return pil

    def _render_page(self):
        if not self.doc or self.page_count == 0:
            self.page_canvas.delete("all")
            return
        if self.fit_width:
            cw = max(100, self.page_canvas.winfo_width() - 20)
            img = self._render_page_image(self.current, scale_by_width=cw)
        else:
            img = self._render_page_image(self.current, scale=ZOOM_STEPS[self.zoom_idx])

        self.page_img_ref = ImageTk.PhotoImage(img)
        self.page_canvas.delete("all")
        # 居中放置
        cw = self.page_canvas.winfo_width()
        ch = self.page_canvas.winfo_height()
        x = max(10, (cw - img.width) // 2)
        y = max(10, (ch - img.height) // 2)
        # 阴影
        self.page_canvas.create_rectangle(
            x + 3, y + 3, x + img.width + 3, y + img.height + 3,
            fill="#2a2a2a", outline=""
        )
        self.page_canvas.create_image(x, y, anchor="nw", image=self.page_img_ref)
        sw = max(cw, img.width + 40)
        sh = max(ch, img.height + 40)
        self.page_canvas.configure(scrollregion=(0, 0, sw, sh))

    def _on_canvas_resize(self, _evt):
        if self._resize_after:
            self.after_cancel(self._resize_after)
        self._resize_after = self.after(80, self._render_page)

    def _on_wheel(self, evt):
        if not self.doc:
            return
        direction = -1 if evt.delta > 0 else 1  # 向下滑 = 翻下一页
        top, bottom = self.page_canvas.yview()
        at_top = top <= 0.0001
        at_bottom = bottom >= 0.9999
        fits = at_top and at_bottom  # 整页已完全显示

        if fits:
            # 页面完全可见，直接翻页
            if direction > 0:
                self.next_page()
            else:
                self.prev_page()
            return

        # 滚到边缘再翻页
        if direction > 0 and at_bottom:
            if self.current < self.page_count - 1:
                self.goto(self.current + 1)
                self.page_canvas.yview_moveto(0.0)
            return
        if direction < 0 and at_top:
            if self.current > 0:
                self.goto(self.current - 1)
                self.page_canvas.yview_moveto(1.0)
            return

        self.page_canvas.yview_scroll(int(-evt.delta / 120), "units")

    def _on_ctrl_wheel(self, evt):
        if evt.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()

    # ----- 导航 -----
    def goto(self, idx):
        if not self.doc:
            return
        idx = max(0, min(self.page_count - 1, idx))
        self.current = idx
        self._render_page()
        self._highlight_current_thumb()
        self._update_status()

    def prev_page(self):
        self.goto(self.current - 1)

    def next_page(self):
        self.goto(self.current + 1)

    def on_page_entry(self, _evt=None):
        try:
            n = int(self.page_var.get().split("/")[0].strip())
            self.goto(n - 1)
        except Exception:
            self._update_status()

    def on_thumb_click(self, evt):
        y = self.thumb_canvas.canvasy(evt.y)
        x = self.thumb_canvas.canvasx(evt.x)
        for i, (_rect, _item, _label, w, h) in enumerate(self.thumb_items):
            rect_coords = self.thumb_canvas.coords(self.thumb_items[i][0])
            if rect_coords[1] <= y <= rect_coords[3]:
                self.goto(i)
                return

    # ----- 缩放 / 旋转 -----
    def zoom_in(self):
        self.fit_width = False
        if self.zoom_idx < len(ZOOM_STEPS) - 1:
            self.zoom_idx += 1
        self._render_page()
        self._update_status()

    def zoom_out(self):
        self.fit_width = False
        if self.zoom_idx > 0:
            self.zoom_idx -= 1
        self._render_page()
        self._update_status()

    def toggle_fit(self):
        self.fit_width = not self.fit_width
        self._render_page()
        self._update_status()

    def rotate_current(self, delta):
        if not self.doc:
            return
        cur = self.rotations.get(self.current, 0)
        self.rotations[self.current] = (cur + delta) % 360
        if self.rotations[self.current] == 0:
            self.rotations.pop(self.current, None)
        # 重渲染当前页 + 更新缩略图
        self._render_page()
        self._refresh_single_thumb(self.current)
        self._update_status()

    def _refresh_single_thumb(self, idx):
        # 简化：整体重建缩略图（数量少时够用）
        self._render_thumbnails()

    # ----- 状态栏 -----
    def _update_status(self):
        if not self.doc:
            self.page_var.set("0 / 0")
            self.zoom_var.set("100%")
            self.status.set("未打开文件")
            return
        self.page_var.set(f"{self.current + 1} / {self.page_count}")
        if self.fit_width:
            self.zoom_var.set("适合")
        else:
            self.zoom_var.set(f"{int(ZOOM_STEPS[self.zoom_idx] * 100)}%")
        dirty = "  ● 有未保存的修改" if self.rotations else ""
        name = os.path.basename(self.path) if self.path else ""
        self.status.set(f"{name}    第 {self.current + 1} 页 / 共 {self.page_count} 页{dirty}")


# ---------- 拆分 ----------

class SplitTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.input_path = tk.StringVar()
        self.out_dir = tk.StringVar()
        self.ranges = tk.StringVar()

        row = 0
        ttk.Label(self, text="输入 PDF：").grid(row=row, column=0, sticky="w")
        ttk.Entry(self, textvariable=self.input_path, width=50).grid(row=row, column=1, sticky="ew")
        ttk.Button(self, text="选择…", command=self.pick_input).grid(row=row, column=2, padx=4)

        row += 1
        ttk.Label(self, text="输出目录：").grid(row=row, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(self, textvariable=self.out_dir, width=50).grid(row=row, column=1, sticky="ew", pady=(6, 0))
        ttk.Button(self, text="选择…", command=self.pick_out_dir).grid(row=row, column=2, padx=4, pady=(6, 0))

        row += 1
        ttk.Label(self, text="页码规则：").grid(row=row, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(self, textvariable=self.ranges, width=50).grid(row=row, column=1, sticky="ew", pady=(6, 0))

        row += 1
        hint = (
            "留空 = 每页一个文件；\n"
            "“1-3,5” = 1~3 和 5 页合为一个文件；\n"
            "“1-3;4-6” = 输出两个文件。"
        )
        ttk.Label(self, text=hint, foreground="#666").grid(
            row=row, column=1, sticky="w", pady=(4, 8)
        )

        row += 1
        ttk.Button(self, text="开始拆分", command=self.run).grid(row=row, column=1, sticky="e")

        self.columnconfigure(1, weight=1)

        register_drop(self, self._on_drop)

    def _on_drop(self, paths):
        for p in paths:
            if p.lower().endswith(".pdf") and os.path.isfile(p):
                self.input_path.set(p)
                if not self.out_dir.get():
                    self.out_dir.set(os.path.dirname(p))
                return

    def pick_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if path:
            self.input_path.set(path)
            if not self.out_dir.get():
                self.out_dir.set(os.path.dirname(path))

    def pick_out_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.out_dir.set(d)

    def run(self):
        src = self.input_path.get().strip()
        out = self.out_dir.get().strip()
        if not src or not os.path.isfile(src):
            messagebox.showerror("错误", "请选择有效的输入 PDF")
            return
        if not out:
            messagebox.showerror("错误", "请选择输出目录")
            return
        os.makedirs(out, exist_ok=True)

        try:
            reader = PdfReader(src)
            groups = parse_page_ranges(self.ranges.get(), len(reader.pages))
        except Exception as e:
            messagebox.showerror("错误", str(e))
            return

        base = os.path.splitext(os.path.basename(src))[0]
        count = 0
        for i, pages in enumerate(groups, 1):
            writer = PdfWriter()
            for p in pages:
                writer.add_page(reader.pages[p])
            if len(pages) == 1:
                name = f"{base}_p{pages[0] + 1}.pdf"
            else:
                name = f"{base}_part{i}.pdf"
            with open(os.path.join(out, name), "wb") as f:
                writer.write(f)
            count += 1
        messagebox.showinfo("完成", f"已生成 {count} 个文件到：\n{out}")


# ---------- 合并 ----------

class MergeTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.files = []

        hint = "待合并文件（支持 PDF 和图片 png/jpg/bmp/tif/gif/webp，可拖拽添加）："
        ttk.Label(self, text=hint).grid(row=0, column=0, columnspan=2, sticky="w")

        list_wrap = ttk.Frame(self)
        list_wrap.grid(row=1, column=0, sticky="nsew", pady=4)
        self.listbox = tk.Listbox(list_wrap, height=14, selectmode=tk.EXTENDED, activestyle="dotbox")
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        btns = ttk.Frame(self)
        btns.grid(row=1, column=1, sticky="ns", padx=6)
        ttk.Button(btns, text="添加…", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(btns, text="上移", command=lambda: self.move(-1)).pack(fill="x", pady=2)
        ttk.Button(btns, text="下移", command=lambda: self.move(1)).pack(fill="x", pady=2)
        ttk.Button(btns, text="移除", command=self.remove).pack(fill="x", pady=2)
        ttk.Button(btns, text="清空", command=self.clear).pack(fill="x", pady=2)

        ttk.Button(self, text="合并并保存…", command=self.run).grid(
            row=2, column=0, columnspan=2, sticky="e", pady=(8, 0)
        )

        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        register_drop(self, self._on_drop)
        register_drop(self.listbox, self._on_drop)

    def _on_drop(self, paths):
        accepted = 0
        for p in paths:
            ext = os.path.splitext(p)[1].lower()
            if os.path.isfile(p) and (ext == ".pdf" or ext in IMAGE_EXTS):
                self._add_one(p)
                accepted += 1
        if accepted == 0:
            messagebox.showwarning("提示", "仅支持 PDF 或图片文件")

    def _add_one(self, p):
        self.files.append(p)
        label = f"{'🖼 ' if os.path.splitext(p)[1].lower() in IMAGE_EXTS else '📄 '}{p}"
        self.listbox.insert(tk.END, label)

    def add_files(self):
        types = [
            ("PDF / 图片", "*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.gif *.webp"),
            ("PDF", "*.pdf"),
            ("图片", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.gif *.webp"),
            ("所有文件", "*.*"),
        ]
        paths = filedialog.askopenfilenames(filetypes=types)
        for p in paths:
            self._add_one(p)

    def remove(self):
        for idx in sorted(self.listbox.curselection(), reverse=True):
            del self.files[idx]
            self.listbox.delete(idx)

    def clear(self):
        self.files.clear()
        self.listbox.delete(0, tk.END)

    def move(self, delta):
        sel = list(self.listbox.curselection())
        if not sel:
            return
        if delta < 0 and sel[0] == 0:
            return
        if delta > 0 and sel[-1] == len(self.files) - 1:
            return
        for idx in sel if delta < 0 else reversed(sel):
            new = idx + delta
            self.files[idx], self.files[new] = self.files[new], self.files[idx]
            text = self.listbox.get(idx)
            self.listbox.delete(idx)
            self.listbox.insert(new, text)
        self.listbox.selection_clear(0, tk.END)
        for idx in sel:
            self.listbox.selection_set(idx + delta)

    def run(self):
        if len(self.files) < 2:
            messagebox.showerror("错误", "至少选择两个文件")
            return
        out = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="merged.pdf",
        )
        if not out:
            return
        try:
            writer = PdfWriter()
            for path in self.files:
                ext = os.path.splitext(path)[1].lower()
                if ext in IMAGE_EXTS:
                    data = image_to_pdf_bytes(path)
                    reader = PdfReader(io.BytesIO(data))
                else:
                    reader = PdfReader(path)
                for page in reader.pages:
                    writer.add_page(page)
            with open(out, "wb") as f:
                writer.write(f)
        except Exception as e:
            messagebox.showerror("错误", str(e))
            return
        messagebox.showinfo("完成", f"已合并为：\n{out}")


# ---------- 入口 ----------

def main():
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    root.title("PDF 工具")
    root.geometry("1100x720")

    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True)
    viewer = ViewerTab(nb)
    nb.add(viewer, text="查看 / 编辑")
    nb.add(SplitTab(nb), text="拆分")
    nb.add(MergeTab(nb), text="合并（PDF + 图片）")

    root.mainloop()


if __name__ == "__main__":
    main()
