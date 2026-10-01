import ctypes
import os
import queue
import threading
import traceback
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pythoncom

import ar
import generate

FONT = "맑은 고딕"
# Basic-info boxes: (pale background, title color) so neighbouring boxes are easy to tell apart.
SECTIONS = {"감사인": ("#E6F0FA", "#2F5A8A"), "날짜": ("#E6F4EA", "#2E6B45"), "회사": ("#FAF0E2", "#8A5A1F")}


class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.vars = {k: tk.StringVar() for group in generate.FIELDS.values() for k, _ in group}
        for spec in generate.FORMS.values():
            for tok, _, _ in spec["fields"]:
                self.vars[f"{spec['tab']}.{tok}"] = tk.StringVar()
        self.list_path = tk.StringVar()
        self.ar_path = tk.StringVar()
        self.kinds = []
        root.title(f"금융거래조회서 생성 v{generate.VERSION}")
        root.protocol("WM_DELETE_WINDOW", self.close)
        scale = root.winfo_fpixels("1i") / 96
        self.px = lambda n: int(n * scale)
        for name in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont"):
            tkfont.nametofont(name).configure(family=FONT, size=11)
        style = ttk.Style()
        style.configure("Treeview", rowheight=self.px(26))
        style.configure("TNotebook.Tab", padding=(self.px(12), self.px(4)))
        for name, (bg, fg) in SECTIONS.items():
            style.configure(f"{name}.TLabelframe", background=bg)
            style.configure(f"{name}.TLabelframe.Label", foreground=fg, font=(FONT, 11, "bold"))
            style.configure(f"{name}.TLabel", background=bg)
            style.configure(f"{name}.TFrame", background=bg)

        f = ttk.Frame(root, padding=10)
        f.pack(fill="both", expand=True)
        f.columnconfigure(0, weight=1)
        f.rowconfigure(2, weight=1)

        self.nb = ttk.Notebook(f)
        self.nb.grid(row=0, column=0, sticky="nsew")
        self.basic_tab()
        self.form_tabs = {kind: self.form_tab(kind, spec) for kind, spec in generate.FORMS.items()}
        self.ar = self.ar_tab()
        self.nb.bind("<<NotebookTabChanged>>", self.on_tab)

        # Financial-list bar; hidden on the 채권채무 tab, which has its own file and button.
        self.bar = bar = ttk.Frame(f, padding=(0, 8))
        bar.grid(row=1, column=0, sticky="we")
        bar.columnconfigure(1, weight=1)
        ttk.Label(bar, text="금융기관 조회처 목록").grid(row=0, column=0, padx=(0, 8))
        ttk.Entry(bar, textvariable=self.list_path).grid(row=0, column=1, sticky="we")
        ttk.Button(bar, text="찾아보기", command=self.pick).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(bar, text="다시 읽기", command=self.refresh).grid(row=0, column=3, padx=(6, 0))
        self.btn = ttk.Button(bar, text="금융거래조회서 생성", command=self.start_fin, width=18)
        self.btn.grid(row=0, column=4, padx=(16, 0), ipady=4)

        self.log = scrolledtext.ScrolledText(f, height=7, state="disabled", font=(FONT, 10))
        self.log.grid(row=2, column=0, sticky="nsew")
        ttk.Label(f, text=f"v{generate.VERSION} · 제작: apc25", foreground="gray", font=(FONT, 9)).grid(row=3, column=0, sticky="e",
                                                                              pady=(4, 0))

        self.load()
        self.vars["조회기준일"].trace_add("write", self.recalc)
        self.refresh()
        self.refresh_ar()
        root.after(100, self.poll)

    def entry_rows(self, box, fields, section=None):
        widgets = {}
        for i, (key, label, kind) in enumerate(fields):
            ttk.Label(box, text=label, style=f"{section}.TLabel" if section else "TLabel").grid(
                row=i, column=0, sticky="w", padx=(0, 8), pady=2)
            if kind == "yn":
                w = ttk.Combobox(box, textvariable=self.vars[key], values=("여", "부"), state="readonly", width=6)
                w.grid(row=i, column=1, sticky="w", pady=2)
            else:
                w = ttk.Entry(box, textvariable=self.vars[key], width=28)
                w.grid(row=i, column=1, sticky="we", pady=2)
            widgets[key] = (w, kind)
        return widgets

    def basic_tab(self):
        tab = ttk.Frame(self.nb, padding=8)
        self.nb.add(tab, text="기본정보")
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)

        def common(group):
            return [(k, lab, "text") for k, lab in generate.FIELDS[group]]

        def section(name, **grid):
            box = ttk.LabelFrame(tab, text=name, padding=8, style=f"{name}.TLabelframe")
            box.grid(sticky="nsew", padx=4, pady=4, **grid)
            box.columnconfigure(1, weight=1)
            self.entry_rows(box, common(name), section=name)
            return box

        section("감사인", row=0, column=0)
        box = section("날짜", row=1, column=0)
        ttk.Label(box, text="날짜는 2026-12-31 형식. 조회기준일을 바꾸면 유효기간과 양식 탭의 기간이 다시 계산됩니다.",
                  foreground="gray", wraplength=self.px(470), style="날짜.TLabel").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))
        box = section("회사", row=0, column=1, rowspan=2)
        self.biz_box(box, len(generate.FIELDS["회사"]))

    def biz_box(self, parent, start):
        ttk.Label(parent, text="사업자등록번호", style="회사.TLabel").grid(row=start, column=0, sticky="nw", pady=(8, 2))
        self.biz = ttk.Treeview(parent, columns=("no", "name"), show="headings", height=4)
        self.biz.heading("no", text="사업자등록번호")
        self.biz.heading("name", text="사업장명")
        self.biz.column("no", width=self.px(150))
        self.biz.column("name", width=self.px(170))
        self.biz.grid(row=start, column=1, sticky="we", pady=(8, 2))
        add = ttk.Frame(parent, style="회사.TFrame")
        add.grid(row=start + 1, column=1, sticky="we")
        self.biz_no, self.biz_name = tk.StringVar(), tk.StringVar()
        ttk.Label(add, text="번호", style="회사.TLabel").pack(side="left", padx=(0, 4))
        ttk.Entry(add, textvariable=self.biz_no, width=13).pack(side="left")
        ttk.Label(add, text="사업장명", style="회사.TLabel").pack(side="left", padx=(8, 4))
        ttk.Entry(add, textvariable=self.biz_name, width=9).pack(side="left")
        ttk.Button(add, text="추가", width=5, command=self.biz_add).pack(side="left", padx=(6, 0))
        ttk.Button(add, text="삭제", width=5, command=self.biz_del).pack(side="left", padx=(4, 0))
        ttk.Label(parent, text="번호와 사업장명을 넣고 [추가]. "
                               "2개 이상이면 요구서에 '별도 첨부'로 적고, 목록 쪽을 조회서 끝에 붙입니다.",
                  foreground="gray", wraplength=self.px(470), style="회사.TLabel").grid(
            row=start + 2, column=0, columnspan=2, sticky="w", pady=(6, 0))

    def form_tab(self, kind, spec):
        tab = ttk.Frame(self.nb, padding=8)
        self.nb.add(tab, text=spec["tab"])
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)
        box = ttk.LabelFrame(tab, text="이 양식에만 필요한 입력", padding=8)
        box.grid(row=0, column=0, sticky="we", padx=4, pady=4)
        fields = [(f"{spec['tab']}.{tok}", lab, typ) for tok, lab, typ in spec["fields"]]
        widgets = self.entry_rows(box, fields)
        ttk.Label(box, text=spec["note"], foreground="gray").grid(row=len(fields), column=0, columnspan=2,
                                                                  sticky="w", pady=(4, 0))
        status = ttk.Label(tab, text="")
        status.grid(row=1, column=0, sticky="w", padx=4, pady=(8, 2))
        tree = ttk.Treeview(tab, columns=("no", "to", "route", "rows", "acct"), show="headings", height=6)
        for col, text, width in (("no", "순번", 50), ("to", "수신처", 220), ("route", "구분", 160),
                                 ("rows", "목록 행", 80), ("acct", "계정과목", 360)):
            tree.heading(col, text=text)
            tree.column(col, width=self.px(width), anchor="center" if col in ("no", "rows") else "w")
        tree.grid(row=2, column=0, sticky="nsew", padx=4)
        return {"tab": tab, "widgets": widgets, "status": status, "tree": tree}

    def ar_tab(self):
        tab = ttk.Frame(self.nb, padding=8)
        self.nb.add(tab, text="채권채무")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)
        box = ttk.LabelFrame(tab, text="거래처 입력 파일", padding=8)
        box.grid(row=0, column=0, sticky="we", padx=4, pady=4)
        box.columnconfigure(0, weight=1)
        ttk.Entry(box, textvariable=self.ar_path).grid(row=0, column=0, sticky="we")
        ttk.Button(box, text="찾아보기", command=self.pick_ar).grid(row=0, column=1, padx=(6, 0))
        ttk.Button(box, text="다시 읽기", command=self.refresh_ar).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(box, text="빈 입력 양식 만들기", command=self.new_ar_input).grid(row=0, column=3, padx=(6, 0))
        rows = ar.AR_ROWS
        ttk.Label(box, text="기본정보 탭의 감사인명·담당 공인회계사·전화·회사명·대표자·작성일·조회기준일을 씁니다. "
                            f"한 거래처의 채권 또는 채무가 {rows}줄을 넘으면 표에 {rows - 1}줄과 '기타(별지 참조)'를 넣고 "
                            "전체 내역을 별지로 붙입니다.",
                  foreground="gray", wraplength=self.px(880)).grid(row=1, column=0, columnspan=4, sticky="w",
                                                                    pady=(6, 0))
        self.ar_status = ttk.Label(tab, text="")
        self.ar_status.grid(row=1, column=0, sticky="w", padx=4, pady=(8, 2))
        self.ar_tree = ttk.Treeview(tab, columns=("i", "no", "name", "ar", "ap", "n", "extra"), show="headings",
                                    height=6)
        for col, text, width in (("i", "순번", 50), ("no", "조회서번호", 110), ("name", "거래처명", 260),
                                 ("ar", "채권 합계", 150), ("ap", "채무 합계", 150), ("n", "줄 수", 70),
                                 ("extra", "별지", 70)):
            self.ar_tree.heading(col, text=text)
            self.ar_tree.column(col, width=self.px(width), anchor="e" if col in ("ar", "ap") else
                                "w" if col == "name" else "center")
        self.ar_tree.grid(row=2, column=0, sticky="nsew", padx=4)
        self.ar_btn = ttk.Button(tab, text="채권채무조회서 생성", command=self.start_ar, width=18)
        self.ar_btn.grid(row=3, column=0, sticky="e", padx=4, pady=(8, 0), ipady=4)
        return tab

    def on_tab(self, _):
        if self.nb.select() == str(self.ar):
            self.bar.grid_remove()
        else:
            self.bar.grid()

    def refresh_ar(self):
        self.ar_tree.delete(*self.ar_tree.get_children())
        path = self.ar_path.get().strip()
        if not path or not Path(path).is_file():
            self.ar_status.configure(text="거래처 입력 파일을 고르세요. 없으면 [빈 입력 양식 만들기]로 만든 뒤 채웁니다.")
            return
        try:
            letters = ar.load_ar(path)
        except Exception as e:
            self.ar_status.configure(text="입력 파일을 확인해야 합니다. 아래 로그를 보세요.")
            self.write(f"[채권채무] {e}")
            return
        for i, ((no, name), lines) in enumerate(letters, 1):
            secs = {sec: [x for x in lines if x["kind"] == sec] for sec in ar.AR_KINDS}
            total = {sec: ar.fmt_amount(sum(x["amount"] for x in v)) for sec, v in secs.items()}
            extra = "별지" if any(len(v) > ar.AR_ROWS for v in secs.values()) else ""
            self.ar_tree.insert("", "end", values=(i, no or f"{i:03d}", name, total["채권"], total["채무"],
                                                   len(lines), extra))
        self.ar_status.configure(text=f"이 파일로 만들 채권채무조회서: {len(letters)}통")

    def pick_ar(self):
        p = filedialog.askopenfilename(title="채권채무 거래처 입력 파일 선택", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.ar_path.set(p)
            self.refresh_ar()

    def new_ar_input(self):
        p = filedialog.asksaveasfilename(title="빈 입력 양식 저장", defaultextension=".xlsx",
                                         initialfile="채권채무_거래처목록.xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            ar.write_ar_input_template(p)
            self.ar_path.set(p)
            self.refresh_ar()
            os.startfile(p)

    def refresh(self):
        path = self.list_path.get().strip()
        letters = []
        if path and Path(path).is_file():
            try:
                letters, skipped = generate.load_letters(path)
                for reason, rows in skipped.items():
                    self.write(f"[목록] 제외 {len(rows)}행 — {reason}: {rows}")
            except Exception as e:
                self.write(f"[목록] 읽기 오류: {e}")
        self.kinds = [k for k in generate.FORMS if any(kind == k for (kind, _), _ in letters)]
        for kind, t in self.form_tabs.items():
            t["tree"].delete(*t["tree"].get_children())
            n = 0
            for i, ((k, to), rows) in enumerate(letters, 1):
                if k != kind:
                    continue
                n += 1
                route = generate.ONLINE_DIR if generate.is_online(rows) else generate.MAIL_DIR
                accts = ", ".join(sorted({str(x["acct"]) for x in rows}))
                t["tree"].insert("", "end", values=(i, to, route, ", ".join(str(x["row"]) for x in rows), accts))
            t["status"].configure(text=f"이번 목록에서 이 양식으로 만들 조회서: {n}통" if n else
                                  "이번 목록에는 이 양식 대상이 없습니다. 위 입력은 비어 있어도 됩니다.")
            for w, typ in t["widgets"].values():
                w.configure(state=("readonly" if typ == "yn" else "normal") if n else "disabled")

    def biz_add(self):
        no = self.biz_no.get().strip()
        if no:
            self.biz.insert("", "end", values=(no, self.biz_name.get().strip()))
            self.biz_no.set("")
            self.biz_name.set("")

    def biz_del(self):
        for item in self.biz.selection():
            self.biz.delete(item)

    def recalc(self, *_):
        try:
            base = generate.parse_date(self.vars["조회기준일"].get(), "조회기준일")
        except generate.InputError:
            return
        for k, v in generate.derived_dates(base).items():
            self.vars[k].set(v)

    def load(self):
        saved = generate.load_saved()
        for k, var in self.vars.items():
            var.set(saved.get(k, ""))
        for no, name in saved.get("사업자목록", []):
            self.biz.insert("", "end", values=(no, name))
        self.list_path.set(saved.get("조회처목록", ""))
        self.ar_path.set(saved.get("채권채무목록", ""))

    def collect(self):
        raw = {k: v.get().strip() for k, v in self.vars.items()}
        raw["사업자목록"] = [list(map(str, self.biz.item(i, "values"))) for i in self.biz.get_children()]
        raw["조회처목록"] = self.list_path.get().strip()
        raw["채권채무목록"] = self.ar_path.get().strip()
        return raw

    def close(self):
        generate.save_values(self.collect())
        self.root.destroy()

    def pick(self):
        p = filedialog.askopenfilename(title="조회처 목록 선택", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.list_path.set(p)
            self.refresh()

    def write(self, line):
        self.log.configure(state="normal")
        self.log.insert("end", line + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start_fin(self):
        raw = self.collect()
        generate.save_values(raw)
        if not Path(raw["조회처목록"]).is_file():
            messagebox.showwarning("조회처 목록", "금융기관 조회처 목록 파일을 먼저 고르세요.")
            return
        self.refresh()
        try:
            generate.prepare(raw, self.kinds)
        except generate.InputError as e:
            messagebox.showwarning("입력 확인", str(e))
            return
        self.begin(generate.run, raw, Path(raw["조회처목록"]))

    def start_ar(self):
        raw = self.collect()
        generate.save_values(raw)
        if not Path(raw["채권채무목록"]).is_file():
            messagebox.showwarning("거래처 입력 파일", "채권채무 거래처 입력 파일을 먼저 고르세요.")
            return
        try:
            ar.prepare_ar(raw)
            ar.load_ar(raw["채권채무목록"])
        except generate.InputError as e:
            messagebox.showwarning("입력 확인", str(e))
            return
        self.begin(ar.run_ar, raw, Path(raw["채권채무목록"]))

    def begin(self, fn, raw, path):
        self.btn.configure(state="disabled")
        self.ar_btn.configure(state="disabled")
        self.write("시작합니다. 한글이 열리고 '접근 허용' 팝업이 뜨면 '모두 허용'을 누르세요.")
        threading.Thread(target=self.work, args=(fn, raw, path), daemon=True).start()

    def work(self, fn, raw, path):
        pythoncom.CoInitialize()
        try:
            self.q.put(("done", fn(raw, path, log=self.q.put)))
        except generate.InputError as e:
            self.q.put(("error", str(e)))
        except Exception:
            self.q.put(("error", traceback.format_exc()))
        finally:
            pythoncom.CoUninitialize()

    def poll(self):
        while not self.q.empty():
            msg = self.q.get()
            if isinstance(msg, str):
                self.write(msg)
                continue
            kind, payload = msg
            self.btn.configure(state="normal")
            self.ar_btn.configure(state="normal")
            if kind == "done":
                self.write("완료. 입력값과 생성 목록은 결과 폴더의 생성기록.xlsx에 남겼습니다.")
                os.startfile(payload)
            else:
                self.write("오류: " + payload)
                messagebox.showerror("오류", payload)
        self.root.after(100, self.poll)


if __name__ == "__main__":
    # Render at the monitor's real DPI; without this Windows bitmap-stretches the window and text looks blurry.
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    root = tk.Tk()
    App(root)
    root.mainloop()
