import json
import os
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox, colorchooser
import tkinter.font as tkfont
from dataclasses import dataclass, asdict
from datetime import datetime
import webbrowser
import html as htmlmod

APP_TITLE = "OwnYourWords with No-Subscription"
DEFAULT_FONT = "Segoe UI"
DEFAULT_SIZE = 12

INDENT_STEP_PX = 28
MAX_INDENT_LEVEL = 12

PAGE_BREAK_TOKEN = "<<PAGE_BREAK>>"


@dataclass
class Comment:
    id: str
    start: str
    end: str
    text: str
    created_at: str


class WordLite(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x800")

        self.current_file = None
        self.comments: list[Comment] = []
        self.comment_counter = 0

        # toolbar relayout debounce state
        self._tb_widgets: list[tk.Widget] = []
        self._tb_relayout_job = None
        self._tb_last_width = None

        self._build_ui()
        self._apply_default_style()

    # ---------------- Toolbar (wrapping, debounced) ----------------
    def _make_wrapping_toolbar(self):
        bar = ttk.Frame(self, padding=(8, 6))
        bar.pack(side=tk.TOP, fill=tk.X)

        inner = ttk.Frame(bar)
        inner.pack(fill=tk.X)

        self._tb_inner = inner

        def schedule_relayout(event=None):
            if self._tb_relayout_job is not None:
                try:
                    self.after_cancel(self._tb_relayout_job)
                except Exception:
                    pass
            self._tb_relayout_job = self.after(60, relayout)

        def relayout():
            self._tb_relayout_job = None
            w = inner.winfo_width()
            if w <= 1:
                return
            if self._tb_last_width == w:
                return
            self._tb_last_width = w

            PAD = 12
            x = 0
            row = 0
            col = 0

            for widget in self._tb_widgets:
                widget.grid_forget()

            for widget in self._tb_widgets:
                widget.update_idletasks()
                ww = widget.winfo_reqwidth() + PAD
                if x + ww > w and col > 0:
                    row += 1
                    col = 0
                    x = 0
                widget.grid(row=row, column=col, padx=3, pady=2, sticky="w")
                x += ww
                col += 1

        inner.bind("<Configure>", schedule_relayout)

        def add(widget: tk.Widget):
            self._tb_widgets.append(widget)
            schedule_relayout()

        self._toolbar_add = add

    # ---------------- UI ----------------
    def _build_ui(self):
        self._make_wrapping_toolbar()

        def v_sep():
            self._toolbar_add(ttk.Label(self._tb_inner, text="|"))

        # ---- File ----
        self._toolbar_add(ttk.Button(self._tb_inner, text="New", command=self.new_doc))
        self._toolbar_add(ttk.Button(self._tb_inner, text="Open", command=self.open_doc))
        self._toolbar_add(ttk.Button(self._tb_inner, text="Save", command=self.save_doc))
        self._toolbar_add(ttk.Button(self._tb_inner, text="Save As", command=self.save_as_doc))
        v_sep()

        # ---- Export ----
        self._toolbar_add(ttk.Button(self._tb_inner, text="Export PDF", command=self.export_pdf))
        v_sep()

        # ---- Font + size ----
        self._toolbar_add(ttk.Label(self._tb_inner, text="Font:"))
        self.font_var = tk.StringVar(value=DEFAULT_FONT)
        self.font_box = ttk.Combobox(self._tb_inner, textvariable=self.font_var, width=22, state="readonly")
        self.font_box["values"] = sorted(tkfont.families())
        self.font_box.bind("<<ComboboxSelected>>", lambda e: self.apply_font_to_selection())
        self._toolbar_add(self.font_box)

        self._toolbar_add(ttk.Label(self._tb_inner, text="Size:"))
        self.size_var = tk.IntVar(value=DEFAULT_SIZE)
        self.size_box = ttk.Combobox(self._tb_inner, textvariable=self.size_var, width=5, state="readonly")
        self.size_box["values"] = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48]
        self.size_box.bind("<<ComboboxSelected>>", lambda e: self.apply_font_to_selection())
        self._toolbar_add(self.size_box)
        v_sep()

        # ---- Basic formatting (FIXED) ----
        self._toolbar_add(ttk.Button(self._tb_inner, text="B", command=self.toggle_bold))
        self._toolbar_add(ttk.Button(self._tb_inner, text="I", command=self.toggle_italic))
        self._toolbar_add(ttk.Button(self._tb_inner, text="U", command=self.toggle_underline))
        v_sep()

        # ---- Text color ----
        self._toolbar_add(ttk.Button(self._tb_inner, text="Text Color", command=self.pick_text_color))
        self.color_swatch = tk.Label(self._tb_inner, width=2, relief="groove")
        self._toolbar_add(self.color_swatch)
        self._set_swatch("#000000")
        v_sep()

        # ---- Alignment ----
        self._toolbar_add(ttk.Button(self._tb_inner, text="⟸", command=lambda: self.apply_alignment("left")))
        self._toolbar_add(ttk.Button(self._tb_inner, text="≡", command=lambda: self.apply_alignment("center")))
        self._toolbar_add(ttk.Button(self._tb_inner, text="⟹", command=lambda: self.apply_alignment("right")))
        v_sep()

        # ---- Bullets + Indent ----
        self._toolbar_add(ttk.Button(self._tb_inner, text="Bullets •", command=self.toggle_bullets))
        self._toolbar_add(ttk.Button(self._tb_inner, text="Indent +", command=lambda: self.change_indent(+1)))
        self._toolbar_add(ttk.Button(self._tb_inner, text="Indent -", command=lambda: self.change_indent(-1)))
        v_sep()

        # ---- Page break ----
        self._toolbar_add(ttk.Button(self._tb_inner, text="Page Break", command=self.insert_page_break))
        v_sep()

        # ---- Headings ----
        self._toolbar_add(ttk.Button(self._tb_inner, text="H1", command=lambda: self.apply_heading(1)))
        self._toolbar_add(ttk.Button(self._tb_inner, text="H2", command=lambda: self.apply_heading(2)))
        v_sep()

        # ---- Comments ----
        self._toolbar_add(ttk.Button(self._tb_inner, text="Add Comment", command=self.add_comment))

        # ---- Main body ----
        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(body, padding=12)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.page = tk.Frame(left, bg="white", highlightbackground="#ccc", highlightthickness=1)
        self.page.pack(fill=tk.BOTH, expand=True)

        self.text = tk.Text(
            self.page,
            wrap="word",
            undo=True,
            padx=34,
            pady=34,
            borderwidth=0,
            highlightthickness=0,
        )
        vs = ttk.Scrollbar(self.page, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=vs.set)
        vs.pack(side=tk.RIGHT, fill=tk.Y)
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ---- Comments pane ----
        right = ttk.Frame(body, width=340, padding=12)
        right.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Label(right, text="Comments", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.comment_list = tk.Listbox(right)
        self.comment_list.pack(fill=tk.BOTH, expand=True, pady=8)
        self.comment_list.bind("<<ListboxSelect>>", lambda e: self.jump_to_comment())

        row = ttk.Frame(right)
        row.pack(fill=tk.X)
        ttk.Button(row, text="Delete", command=self.delete_comment).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="Edit", command=self.edit_comment).pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)

        # Shortcuts
        self.bind_all("<Control-s>", lambda e: self.save_doc())
        self.bind_all("<Control-o>", lambda e: self.open_doc())
        self.bind_all("<Control-b>", lambda e: self.toggle_bold())
        self.bind_all("<Control-i>", lambda e: self.toggle_italic())
        self.bind_all("<Control-u>", lambda e: self.toggle_underline())

    def _apply_default_style(self):
        base = tkfont.Font(family=DEFAULT_FONT, size=DEFAULT_SIZE)
        self.text.configure(font=base)

        # alignment tags
        self.text.tag_configure("align_left", justify="left")
        self.text.tag_configure("align_center", justify="center")
        self.text.tag_configure("align_right", justify="right")

        # headings
        self.text.tag_configure("h1", font=(DEFAULT_FONT, 22, "bold"), spacing1=10, spacing3=10)
        self.text.tag_configure("h2", font=(DEFAULT_FONT, 16, "bold"), spacing1=8, spacing3=8)

        # comment highlighting
        self.text.tag_configure("comment", background="#fff2cc")
        self.text.tag_configure("comment_selected", background="#ffe599")

        self.configure(bg="#f3f3f3")

    # ---------------- Selection helpers ----------------
    def selection(self):
        try:
            return self.text.index("sel.first"), self.text.index("sel.last")
        except tk.TclError:
            return None

    def _selected_line_range(self):
        sel = self.selection()
        if sel:
            start, end = sel
        else:
            start = self.text.index("insert")
            end = self.text.index("insert")

        line_start = self.text.index(f"{start} linestart")
        line_end = self.text.index(f"{end} lineend")
        try:
            line_end_plus = self.text.index(f"{line_end}+1c")
        except tk.TclError:
            line_end_plus = line_end
        return line_start, line_end_plus

    # ---------------- Composite font engine (FIX) ----------------
    def _apply_composite_font(self, start, end, *, toggle=None):
        """
        Tkinter cannot merge font tags (font family/size + bold/italic/underline).
        We apply ONE composite font tag per selection containing the full state.
        """
        family = self.font_var.get()
        size = int(self.size_var.get())

        tags = set(self.text.tag_names(start))

        bold = "style_bold" in tags
        italic = "style_italic" in tags
        underline = "style_underline" in tags

        if toggle == "bold":
            bold = not bold
        elif toggle == "italic":
            italic = not italic
        elif toggle == "underline":
            underline = not underline

        weight = "bold" if bold else "normal"
        slant = "italic" if italic else "roman"
        underline_flag = 1 if underline else 0

        # Remove existing font/style tags across range
        for t in list(self.text.tag_names()):
            if t.startswith("font_"):
                self.text.tag_remove(t, start, end)

        self.text.tag_remove("style_bold", start, end)
        self.text.tag_remove("style_italic", start, end)
        self.text.tag_remove("style_underline", start, end)

        tag = f"font_{family}_{size}_{weight}_{slant}_{underline_flag}".replace(" ", "_")
        self.text.tag_configure(tag, font=(family, size, weight, slant), underline=underline_flag)
        self.text.tag_add(tag, start, end)

        # Re-add markers for detection
        if bold:
            self.text.tag_add("style_bold", start, end)
        if italic:
            self.text.tag_add("style_italic", start, end)
        if underline:
            self.text.tag_add("style_underline", start, end)

    def apply_font_to_selection(self):
        sel = self.selection()
        if not sel:
            return
        self._apply_composite_font(*sel)

    def toggle_bold(self):
        sel = self.selection()
        if not sel:
            return
        self._apply_composite_font(*sel, toggle="bold")

    def toggle_italic(self):
        sel = self.selection()
        if not sel:
            return
        self._apply_composite_font(*sel, toggle="italic")

    def toggle_underline(self):
        sel = self.selection()
        if not sel:
            return
        self._apply_composite_font(*sel, toggle="underline")

    def apply_heading(self, level):
        sel = self.selection()
        if not sel:
            return
        start, end = sel
        self.text.tag_remove("h1", start, end)
        self.text.tag_remove("h2", start, end)
        self.text.tag_add("h1" if level == 1 else "h2", start, end)

    # ---------------- Text color ----------------
    def _set_swatch(self, hex_color: str):
        self.color_swatch.configure(bg=hex_color)

    def pick_text_color(self):
        sel = self.selection()
        if not sel:
            messagebox.showinfo("Text Color", "Select text first.")
            return

        chosen = colorchooser.askcolor(title="Choose text color")
        if not chosen or not chosen[1]:
            return
        hex_color = chosen[1]
        self._set_swatch(hex_color)

        tag = f"color_{hex_color.replace('#', '')}"
        self.text.tag_configure(tag, foreground=hex_color)
        start, end = sel
        self.text.tag_add(tag, start, end)

    # ---------------- Alignment ----------------
    def apply_alignment(self, which: str):
        start, end = self._selected_line_range()
        self.text.tag_remove("align_left", start, end)
        self.text.tag_remove("align_center", start, end)
        self.text.tag_remove("align_right", start, end)
        tag = {"left": "align_left", "center": "align_center", "right": "align_right"}[which]
        self.text.tag_add(tag, start, end)

    # ---------------- Bullets ----------------
    def toggle_bullets(self):
        start, end = self._selected_line_range()

        start_line = int(start.split(".")[0])
        end_line = int(end.split(".")[0])
        if end.endswith(".0") and end_line > start_line:
            end_line -= 1

        all_bulleted = True
        for ln in range(start_line, end_line + 1):
            line_text = self.text.get(f"{ln}.0", f"{ln}.0 lineend")
            if line_text.strip() == "":
                continue
            if not line_text.startswith("• "):
                all_bulleted = False
                break

        for ln in range(start_line, end_line + 1):
            line_text = self.text.get(f"{ln}.0", f"{ln}.0 lineend")
            if line_text.strip() == "":
                continue
            if all_bulleted:
                if line_text.startswith("• "):
                    self.text.delete(f"{ln}.0", f"{ln}.2")
            else:
                if not line_text.startswith("• "):
                    self.text.insert(f"{ln}.0", "• ")

    # ---------------- Indents ----------------
    def _indent_tag_for_level(self, level: int) -> str:
        level = max(0, min(MAX_INDENT_LEVEL, level))
        return f"indent_{level}"

    def _configure_indent_tag(self, level: int):
        level = max(0, min(MAX_INDENT_LEVEL, level))
        px = level * INDENT_STEP_PX
        tag = self._indent_tag_for_level(level)
        self.text.tag_configure(tag, lmargin1=px, lmargin2=px)

    def _current_indent_level(self, index: str) -> int:
        tags = self.text.tag_names(index)
        for t in tags:
            if t.startswith("indent_"):
                try:
                    return int(t.split("_")[1])
                except Exception:
                    pass
        return 0

    def change_indent(self, delta: int):
        start, end = self._selected_line_range()

        start_line = int(start.split(".")[0])
        end_line = int(end.split(".")[0])
        if end.endswith(".0") and end_line > start_line:
            end_line -= 1

        for ln in range(start_line, end_line + 1):
            line_start = f"{ln}.0"
            line_end = f"{ln}.0 lineend"

            for t in self.text.tag_names(line_start):
                if t.startswith("indent_"):
                    self.text.tag_remove(t, line_start, f"{line_end}+1c")

            current = self._current_indent_level(line_start)
            new_level = max(0, min(MAX_INDENT_LEVEL, current + delta))

            tag = self._indent_tag_for_level(new_level)
            self._configure_indent_tag(new_level)
            self.text.tag_add(tag, line_start, f"{line_end}+1c")

    # ---------------- Page breaks ----------------
    def insert_page_break(self):
        idx = self.text.index("insert")
        self.text.insert(idx, "\n" + PAGE_BREAK_TOKEN + "\n")

    # ---------------- Comments ----------------
    def add_comment(self):
        sel = self.selection()
        if not sel:
            messagebox.showinfo("Add Comment", "Select text first, then click Add Comment.")
            return

        text = simpledialog.askstring("Comment", "Add comment:")
        if not text:
            return

        self.comment_counter += 1
        c = Comment(
            id=f"C{self.comment_counter}",
            start=sel[0],
            end=sel[1],
            text=text,
            created_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        )
        self.comments.append(c)
        self.text.tag_add("comment", c.start, c.end)
        self.refresh_comments()

    def refresh_comments(self):
        self.comment_list.delete(0, tk.END)
        for c in self.comments:
            preview = c.text.replace("\n", " ").strip()
            if len(preview) > 45:
                preview = preview[:45] + "…"
            self.comment_list.insert(tk.END, f"{c.id} • {preview}")

    def jump_to_comment(self):
        sel = self.comment_list.curselection()
        if not sel:
            return
        c = self.comments[sel[0]]
        self.text.tag_remove("comment_selected", "1.0", "end")
        self.text.tag_add("comment_selected", c.start, c.end)
        self.text.see(c.start)
        self.text.mark_set("insert", c.start)
        self.text.focus_set()

    def delete_comment(self):
        sel = self.comment_list.curselection()
        if not sel:
            return
        del self.comments[sel[0]]
        self.refresh_comments()

    def edit_comment(self):
        sel = self.comment_list.curselection()
        if not sel:
            return
        c = self.comments[sel[0]]
        new = simpledialog.askstring("Edit", "Edit comment:", initialvalue=c.text)
        if new is not None:
            c.text = new
            self.refresh_comments()

    # ---------------- Save/Open with formatting tags ----------------
    def new_doc(self):
        if not messagebox.askyesno("New", "Discard current document and start a new one?"):
            return
        self.text.delete("1.0", tk.END)
        self.comments.clear()
        self.refresh_comments()
        self.current_file = None
        self.title(APP_TITLE)

    def save_doc(self):
        if not self.current_file:
            return self.save_as_doc()
        self._write_file(self.current_file)

    def save_as_doc(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".wordlite.json",
            filetypes=[("WordLite Documents", "*.wordlite.json"), ("JSON", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        self.current_file = path
        self._write_file(path)
        self.title(f"{APP_TITLE} — {os.path.basename(path)}")

    def _export_tags(self):
        keep_exact = {
            "h1", "h2",
            "align_left", "align_center", "align_right",
            "comment",
            "style_bold", "style_italic", "style_underline",
        }
        keep_prefixes = ("font_", "color_", "indent_")

        exported = []
        for tag in self.text.tag_names():
            if tag in keep_exact or tag.startswith(keep_prefixes):
                ranges = self.text.tag_ranges(tag)
                for i in range(0, len(ranges), 2):
                    exported.append({"tag": tag, "start": str(ranges[i]), "end": str(ranges[i + 1])})
        return exported

    def _import_tags(self, exported):
        for item in exported:
            tag = item["tag"]
            start = item["start"]
            end = item["end"]

            if tag.startswith("font_") and not self.text.tag_cget(tag, "font"):
                # font_FAMILY_SIZE_WEIGHT_SLANT_UNDERLINE
                try:
                    parts = tag.split("_")
                    underline_flag = int(parts[-1])
                    slant = parts[-2]
                    weight = parts[-3]
                    size = int(parts[-4])
                    family = " ".join(parts[1:-4]).replace("  ", " ")
                    self.text.tag_configure(tag, font=(family, size, weight, slant), underline=underline_flag)
                except Exception:
                    pass

            if tag.startswith("color_") and not self.text.tag_cget(tag, "foreground"):
                try:
                    hex_color = "#" + tag.split("_", 1)[1]
                    self.text.tag_configure(tag, foreground=hex_color)
                except Exception:
                    pass

            if tag.startswith("indent_") and not self.text.tag_cget(tag, "lmargin1"):
                try:
                    level = int(tag.split("_")[1])
                    self._configure_indent_tag(level)
                except Exception:
                    pass

            self.text.tag_add(tag, start, end)

    def _write_file(self, path):
        data = {
            "version": 6,
            "text": self.text.get("1.0", "end-1c"),
            "tags": self._export_tags(),
            "comments": [asdict(c) for c in self.comments],
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Saved", f"Saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Save Failed", str(e))

    def open_doc(self):
        path = filedialog.askopenfilename(
            filetypes=[("WordLite Documents", "*.wordlite.json"), ("JSON", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return

        if not messagebox.askyesno("Open", "Discard current document and open another?"):
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.text.delete("1.0", tk.END)
            self.text.insert("1.0", data.get("text", ""))

            for tag in self.text.tag_names():
                if tag != "sel":
                    self.text.tag_remove(tag, "1.0", "end")

            self._import_tags(data.get("tags", []))

            self.comments = [Comment(**c) for c in data.get("comments", [])]
            self.refresh_comments()

            for c in self.comments:
                self.text.tag_add("comment", c.start, c.end)

            self.current_file = path
            self.title(f"{APP_TITLE} — {os.path.basename(path)}")

        except Exception as e:
            messagebox.showerror("Open Failed", str(e))

    # ---------------- Export PDF (via browser print) ----------------
    def export_pdf(self):
        html_path = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML", "*.html"), ("All files", "*.*")]
        )
        if not html_path:
            return

        raw = self.text.get("1.0", "end-1c")
        parts = raw.split(PAGE_BREAK_TOKEN)

        pages_html = []
        for p in parts:
            safe = htmlmod.escape(p).replace("\n", "<br>\n")
            pages_html.append(f'<div class="page">{safe}</div>')

        html = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<title>Export</title>
<style>
  body {{
    background:#f3f3f3;
    margin:0;
    padding:24px;
    font-family:{DEFAULT_FONT}, Arial, sans-serif;
  }}
  .page {{
    background:#fff;
    width:8.5in;
    min-height:11in;
    margin:0 auto 18px auto;
    padding:1in;
    box-shadow:0 2px 10px rgba(0,0,0,0.12);
    font-size:12pt;
    line-height:1.35;
    page-break-after: always;
  }}
  .page:last-child {{
    page-break-after: auto;
  }}
  @media print {{
    body {{ background:#fff; padding:0; }}
    .page {{ box-shadow:none; margin:0; width:auto; min-height:auto; padding:1in; }}
  }}
</style>
</head>
<body>
  {''.join(pages_html)}
</body>
</html>
"""
        try:
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html)
            webbrowser.open(f"file:///{html_path.replace(os.sep, '/')}")
            messagebox.showinfo(
                "Export PDF",
                "Opened export in your browser.\n\nNow press Ctrl+P and choose 'Microsoft Print to PDF' to save."
            )
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))


if __name__ == "__main__":
    app = WordLite()
    app.mainloop()
