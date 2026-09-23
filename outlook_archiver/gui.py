from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from datetime import date, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from .pst import work_root_for_output
from .utils import parse_years, read_json, write_json_atomic


def worker_python() -> str:
    # pythonw has no standard streams; workers need python.exe to report events.
    executable = Path(sys.executable)
    if executable.name.lower() == "pythonw.exe":
        return str(executable.with_name("python.exe"))
    return str(executable)


def completion_title(event: dict[str, Any]) -> str:
    if event.get("scan_stopped") or int(event.get("unprocessed", 0)):
        return "已停止，尚未全部完成"
    if any(int(event.get(key, 0)) for key in ("failed", "partial", "source_unavailable")):
        return "处理结束，有失败或部分完成项"
    if not int(event.get("selected", 0)):
        return "处理结束，没有符合条件的邮件"
    return "转换已完成（有警告，请查看报告）" if int(event.get("warning", 0)) else "转换已完成"


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Outlook / PST 邮件转 Markdown")
        self.geometry("960x700")
        self.minsize(760, 540)
        self.option_add("*Font", ("Microsoft YaHei UI", 10))
        self.catalog: dict[str, Any] = {"stores": [], "folders": []}
        self.visible_folders: list[dict[str, Any]] = []
        self.msg_paths: list[str] = []
        self.pst_work_custom = False
        self.message_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.process: subprocess.Popen[str] | None = None
        self.last_report = ""
        self.last_failures = ""
        self.pending_result: dict[str, Any] | None = None
        runtime = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "OutlookEmailArchiver" / "runtime"
        runtime.mkdir(parents=True, exist_ok=True)
        self.job_config = runtime / "current-job.json"
        self.stop_file = runtime / "stop.request"
        self.folder_result = runtime / "folders.json"
        self._build()
        self._restore_settings()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll_queue)

    def _build(self) -> None:
        self.completion_status = tk.StringVar(value="准备就绪")
        banner = ttk.Frame(self, padding=(10, 8))
        banner.pack(fill="x")
        ttk.Label(banner, textvariable=self.completion_status, wraplength=720,
                  font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, highlightthickness=0)
        page_scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=page_scroll.set)
        page_scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        outer = ttk.Frame(canvas, padding=10)
        page = canvas.create_window((0, 0), window=outer, anchor="nw")
        outer.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(page, width=event.width))
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-event.delta / 120), "units"))
        source = ttk.LabelFrame(outer, text="1. 邮件来源", padding=8)
        source.pack(fill="x")
        self.source_mode = tk.StringVar(value="outlook")
        ttk.Radiobutton(source, text="当前经典 Outlook", variable=self.source_mode, value="outlook", command=self._source_changed).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(source, text="PST 文件（使用完整工作副本）", variable=self.source_mode, value="pst", command=self._source_changed).grid(row=0, column=1, sticky="w", padx=(18, 0))
        ttk.Radiobutton(source, text="MSG / MSD 文件或文件夹", variable=self.source_mode, value="msg", command=self._source_changed).grid(row=0, column=2, sticky="w", padx=(18, 0))
        self.pst_path = tk.StringVar()
        self.pst_entry = ttk.Entry(source, textvariable=self.pst_path, state="disabled")
        self.pst_entry.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        self.pst_button = ttk.Button(source, text="选择 PST 文件", command=self._choose_pst, state="disabled")
        self.pst_button.grid(row=1, column=3, padx=(6, 0), pady=(6, 0))
        self.pst_work_label = ttk.Label(source, text="PST 工作副本：", state="disabled")
        self.pst_work_label.grid(row=4, column=0, sticky="w", pady=(6, 0))
        self.pst_work_root = tk.StringVar()
        self.pst_work_entry = ttk.Entry(source, textvariable=self.pst_work_root, state="disabled")
        self.pst_work_entry.grid(row=4, column=1, columnspan=2, sticky="ew", pady=(6, 0))
        self.pst_work_button = ttk.Button(source, text="选择副本位置", command=self._choose_pst_work_root, state="disabled")
        self.pst_work_button.grid(row=4, column=3, padx=(6, 0), pady=(6, 0))
        self.pst_work_auto_button = ttk.Button(source, text="跟随保存磁盘", command=self._use_default_pst_work_root, state="disabled")
        self.pst_work_auto_button.grid(row=5, column=3, padx=(6, 0), pady=(4, 0))
        self.pst_work_note = ttk.Label(source, text="默认放在保存位置的同一磁盘；完整副本可复用。", state="disabled")
        self.pst_work_note.grid(row=5, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self.load_button = ttk.Button(source, text="读取邮箱和文件夹", command=self._load_folders)
        self.load_button.grid(row=0, column=3, padx=(6, 0))
        self.msg_path_var = tk.StringVar()
        self.msg_entry = ttk.Entry(source, textvariable=self.msg_path_var, state="disabled")
        self.msg_entry.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self.msg_files_button = ttk.Button(source, text="选择 MSG 文件", command=self._choose_msg_files, state="disabled")
        self.msg_files_button.grid(row=2, column=2, padx=(6, 0), pady=(6, 0))
        self.msg_folder_button = ttk.Button(source, text="选择 MSG 文件夹", command=self._choose_msg_folder, state="disabled")
        self.msg_folder_button.grid(row=2, column=3, padx=(6, 0), pady=(6, 0))
        self.msg_clear_button = ttk.Button(source, text="清空 MSG 选择", command=self._clear_msg_paths, state="disabled")
        self.msg_clear_button.grid(row=3, column=3, padx=(6, 0), pady=(4, 0))
        source.columnconfigure(1, weight=1)

        folders = ttk.LabelFrame(outer, text="2. 邮箱与文件夹（可用 Ctrl/Shift 多选）", padding=8)
        folders.pack(fill="x", pady=(7, 0))
        ttk.Label(folders, text="邮箱 / PST：").grid(row=0, column=0, sticky="w")
        self.store_var = tk.StringVar()
        self.store_combo = ttk.Combobox(folders, textvariable=self.store_var, state="readonly")
        self.store_combo.grid(row=0, column=1, sticky="ew")
        self.store_combo.bind("<<ComboboxSelected>>", lambda _event: self._show_folders())
        self.include_subfolders = tk.BooleanVar(value=True)
        ttk.Checkbutton(folders, text="包含子文件夹", variable=self.include_subfolders).grid(row=0, column=2, padx=(8, 0))
        list_frame = ttk.Frame(folders)
        list_frame.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
        self.folder_list = tk.Listbox(list_frame, selectmode=tk.EXTENDED, exportselection=False, height=6)
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.folder_list.yview)
        self.folder_list.configure(yscrollcommand=scroll.set)
        self.folder_list.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        buttons = ttk.Frame(folders)
        buttons.grid(row=2, column=0, columnspan=3, sticky="w", pady=(5, 0))
        ttk.Button(buttons, text="全选当前邮箱", command=lambda: self.folder_list.select_set(0, tk.END)).pack(side="left")
        ttk.Button(buttons, text="清除选择", command=lambda: self.folder_list.selection_clear(0, tk.END)).pack(side="left", padx=6)
        folders.columnconfigure(1, weight=1)

        settings = ttk.LabelFrame(outer, text="3. 时间与保存位置", padding=8)
        settings.pack(fill="x", pady=(8, 0))
        self.date_mode = tk.StringVar(value="all")
        ttk.Radiobutton(settings, text="全部年份", variable=self.date_mode, value="all").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(settings, text="指定年份", variable=self.date_mode, value="years").grid(row=0, column=1, sticky="w")
        self.years = tk.StringVar(value=str(date.today().year))
        ttk.Entry(settings, textvariable=self.years, width=18).grid(row=0, column=2, sticky="w")
        ttk.Label(settings, text="例如：2024, 2025").grid(row=0, column=3, sticky="w", padx=4)
        ttk.Radiobutton(settings, text="自定义日期", variable=self.date_mode, value="range").grid(row=1, column=0, sticky="w", pady=(5, 0))
        self.start_date = tk.StringVar(value=f"{date.today().year}-01-01")
        self.end_date = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(settings, textvariable=self.start_date, width=12).grid(row=1, column=1, sticky="w", pady=(5, 0))
        ttk.Label(settings, text="至").grid(row=1, column=2, sticky="w", pady=(5, 0))
        ttk.Entry(settings, textvariable=self.end_date, width=12).grid(row=1, column=3, sticky="w", pady=(5, 0))
        ttk.Label(settings, text="（结束日期包含当天）").grid(row=1, column=4, sticky="w", pady=(5, 0))
        ttk.Label(settings, text="保存位置：").grid(row=2, column=0, sticky="w", pady=(7, 0))
        self.output_root = tk.StringVar(value=str(Path.home() / "Documents" / "OutlookMarkdownArchive"))
        self.output_root.trace_add("write", self._output_changed)
        ttk.Entry(settings, textvariable=self.output_root).grid(row=2, column=1, columnspan=3, sticky="ew", pady=(7, 0))
        ttk.Button(settings, text="选择保存位置", command=self._choose_output).grid(row=2, column=4, padx=(6, 0), pady=(7, 0))
        self.layout = tk.StringVar(value="folder_first")
        self.index_mode = tk.StringVar(value="incremental")
        ttk.Label(settings, text="索引更新：").grid(row=4, column=0, sticky="w")
        ttk.Radiobutton(settings, text="增量更新（推荐）", variable=self.index_mode, value="incremental").grid(row=4, column=1, columnspan=2, sticky="w")
        ttk.Radiobutton(settings, text="完整重建（较慢）", variable=self.index_mode, value="full").grid(row=4, column=3, sticky="w")
        ttk.Label(settings, text="目录顺序：").grid(row=3, column=0, sticky="w", pady=(5, 0))
        ttk.Radiobutton(settings, text="文件夹优先", variable=self.layout, value="folder_first").grid(row=3, column=1, sticky="w", pady=(5, 0))
        ttk.Radiobutton(settings, text="年份优先", variable=self.layout, value="year_first").grid(row=3, column=2, sticky="w", pady=(5, 0))
        self.include_addresses = tk.BooleanVar(value=True)
        ttk.Checkbutton(settings, text="Markdown 保留可读取的邮件地址", variable=self.include_addresses).grid(row=3, column=3, columnspan=2, sticky="w", pady=(5, 0))
        settings.columnconfigure(3, weight=1)
        self._update_default_pst_work_root()

        content = ttk.LabelFrame(outer, text="4. 导出内容", padding=8)
        content.pack(fill="x", pady=(7, 0))
        ttk.Button(content, text="仅 Markdown", command=lambda: self._set_content_preset(False)).grid(row=0, column=0, sticky="w")
        ttk.Button(content, text="完整归档", command=lambda: self._set_content_preset(True)).grid(row=0, column=1, sticky="w", padx=(6, 14))
        self.save_msg = tk.BooleanVar(value=True)
        self.save_attachments = tk.BooleanVar(value=True)
        self.save_inline_images = tk.BooleanVar(value=True)
        ttk.Checkbutton(content, text="保存 original.msg", variable=self.save_msg, command=self._content_changed).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(content, text="保存普通附件", variable=self.save_attachments, command=self._content_changed).grid(row=0, column=3, sticky="w", padx=(10, 0))
        ttk.Checkbutton(content, text="保存正文图片", variable=self.save_inline_images, command=self._content_changed).grid(row=0, column=4, sticky="w", padx=(10, 0))
        self.content_summary = tk.StringVar()
        ttk.Label(content, textvariable=self.content_summary, foreground="#555555").grid(row=1, column=0, columnspan=5, sticky="w", pady=(5, 0))
        self._content_changed()

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(8, 0))
        self.sample_limit = tk.BooleanVar(value=True)
        ttk.Checkbutton(actions, text="首次样本最多 5 封（通过后取消）", variable=self.sample_limit).pack(side="left", padx=(0, 10))
        self.preview_button = ttk.Button(actions, text="预览", command=lambda: self._run_job("preview"))
        self.preview_button.pack(side="left")
        self.start_button = ttk.Button(actions, text="开始", command=lambda: self._run_job("export"))
        self.start_button.pack(side="left", padx=6)
        self.stop_button = ttk.Button(actions, text="停止", command=self._stop, state="disabled")
        self.stop_button.pack(side="left")
        ttk.Button(actions, text="继续", command=lambda: self._run_job("export")).pack(side="left", padx=6)
        ttk.Button(actions, text="打开输出目录", command=self._open_output).pack(side="right")
        ttk.Button(actions, text="查看结果报告", command=self._open_report).pack(side="right", padx=6)
        ttk.Button(actions, text="查看失败清单", command=self._open_failures).pack(side="right")

        self.progress = ttk.Progressbar(outer, mode="determinate")
        self.progress.pack(fill="x", pady=(8, 0))
        self.counter = tk.StringVar(value="成功 0　跳过 0　警告 0　失败 0")
        ttk.Label(outer, textvariable=self.counter).pack(anchor="w", pady=(4, 0))
        self.status = tk.StringVar(value="请先读取邮箱和文件夹。")
        ttk.Label(outer, textvariable=self.status).pack(anchor="w")
        self.log = tk.Text(outer, height=5, wrap="word", state="disabled")
        self.log.pack(fill="x", pady=(5, 0))

    def _set_content_preset(self, full: bool) -> None:
        self.save_msg.set(full)
        self.save_attachments.set(full)
        self.save_inline_images.set(full)
        self._content_changed()

    def _content_changed(self) -> None:
        selected = [
            label for enabled, label in (
                (self.save_msg.get(), "MSG"),
                (self.save_attachments.get(), "普通附件"),
                (self.save_inline_images.get(), "正文图片"),
            ) if enabled
        ]
        if not selected:
            self.content_summary.set("仅 Markdown：每封邮件直接生成一个日期开头的 .md 文件，不建立单封邮件目录。")
        elif len(selected) == 3:
            self.content_summary.set("完整归档：保存 Markdown、MSG、附件、正文图片和必要的正文源文件。")
        else:
            self.content_summary.set("自定义归档：保存 Markdown + " + "、".join(selected) + "。")

    def _source_changed(self) -> None:
        pst = self.source_mode.get() == "pst"
        msg = self.source_mode.get() == "msg"
        self.pst_entry.configure(state="normal" if pst else "disabled")
        self.pst_button.configure(state="normal" if pst else "disabled")
        pst_state = "normal" if pst else "disabled"
        self.pst_work_label.configure(state=pst_state)
        self.pst_work_entry.configure(state="readonly" if pst else "disabled")
        self.pst_work_button.configure(state=pst_state)
        self.pst_work_auto_button.configure(state=pst_state)
        self.pst_work_note.configure(state=pst_state)
        if pst and not self.pst_work_custom:
            self._update_default_pst_work_root()
        self.msg_entry.configure(state="readonly" if msg else "disabled")
        self.msg_files_button.configure(state="normal" if msg else "disabled")
        self.msg_folder_button.configure(state="normal" if msg else "disabled")
        self.msg_clear_button.configure(state="normal" if msg else "disabled")
        self.load_button.configure(state="disabled" if msg else "normal")
        self.store_combo.configure(state="disabled" if msg else "readonly")
        self.folder_list.configure(state="disabled" if msg else "normal")
        self.catalog = {"stores": [], "folders": []}; self.store_combo["values"] = (); self.folder_list.delete(0, tk.END)

    def _restore_settings(self) -> None:
        if not self.job_config.exists():
            return
        try:
            config = read_json(self.job_config)
            self.index_mode.set(config.get("index_mode", "incremental"))
            if config.get("source_mode") != "msg" and not self.folder_result.exists():
                return
            self.catalog = read_json(self.folder_result) if self.folder_result.exists() else {"stores": [], "folders": []}
            self.source_mode.set(config.get("source_mode", "outlook")); self._source_changed()
            self.pst_path.set(config.get("pst_path", "")); self.output_root.set(config.get("output_root", self.output_root.get()))
            self.pst_work_custom = not bool(config.get("pst_work_root_auto", not config.get("pst_work_root")))
            if self.pst_work_custom and config.get("pst_work_root"):
                self.pst_work_root.set(config["pst_work_root"])
            else:
                self._update_default_pst_work_root()
            self.msg_paths = list(config.get("msg_paths", [])); self.msg_path_var.set("；".join(self.msg_paths))
            self.layout.set(config.get("layout", "folder_first")); self.include_subfolders.set(bool(config.get("include_subfolders", True)))
            self.include_addresses.set(bool(config.get("include_addresses", True)))
            self.save_msg.set(bool(config.get("save_msg", True)))
            self.save_attachments.set(bool(config.get("save_attachments", True)))
            self.save_inline_images.set(bool(config.get("save_inline_images", True)))
            self._content_changed()
            self.sample_limit.set(bool(config.get("test_limit", 5)))
            # _source_changed 清空了 catalog，重新载入已缓存的只读文件夹目录。
            self.catalog = read_json(self.folder_result) if self.source_mode.get() != "msg" else {"stores": [], "folders": []}
            values = [item["store_name"] + (f"（{item['store_path']}）" if item.get("store_path") else "") for item in self.catalog.get("stores", [])]
            self.store_combo["values"] = values
            store_index = next((i for i, item in enumerate(self.catalog.get("stores", [])) if item["store_id"] == config.get("store_id")), -1)
            if store_index >= 0:
                self.store_combo.current(store_index); self._show_folders()
                wanted = set(config.get("folder_ids", []))
                for index, folder in enumerate(self.visible_folders):
                    if folder["folder_id"] in wanted: self.folder_list.select_set(index)
            if config.get("years"):
                self.date_mode.set("years"); self.years.set(", ".join(str(v) for v in config["years"]))
            elif config.get("start_date") and config.get("end_date"):
                self.date_mode.set("range"); self.start_date.set(config["start_date"])
                self.end_date.set((date.fromisoformat(config["end_date"]) - timedelta(days=1)).isoformat())
            self.status.set("已恢复上次选择；可点击“继续”，程序会核验并跳过已完成邮件。")
        except Exception:
            self.catalog = {"stores": [], "folders": []}

    def _choose_pst(self) -> None:
        value = filedialog.askopenfilename(title="选择 PST 邮件备份", filetypes=[("Outlook 数据文件", "*.pst")])
        if value:
            self.pst_path.set(value)

    def _choose_pst_work_root(self) -> None:
        value = filedialog.askdirectory(title="选择 PST 工作副本位置（需要容纳完整 PST）")
        if value:
            self.pst_work_custom = True
            self.pst_work_root.set(value)

    def _use_default_pst_work_root(self) -> None:
        self.pst_work_custom = False
        self._update_default_pst_work_root()

    def _output_changed(self, *_: object) -> None:
        if not self.pst_work_custom:
            self._update_default_pst_work_root()

    def _update_default_pst_work_root(self) -> None:
        output = self.output_root.get().strip() if hasattr(self, "output_root") else ""
        if output:
            self.pst_work_root.set(str(work_root_for_output(output)))

    def _choose_msg_files(self) -> None:
        values = filedialog.askopenfilenames(title="选择一个或多个 MSG 文件", filetypes=[("Outlook 邮件", "*.msg *.msd"), ("MSG 文件", "*.msg"), ("MSD 兼容尝试", "*.msd")])
        if values:
            self.msg_paths = list(dict.fromkeys([*self.msg_paths, *values]))
            self.msg_path_var.set("；".join(self.msg_paths))

    def _choose_msg_folder(self) -> None:
        value = filedialog.askdirectory(title="选择包含 MSG 的文件夹（会递归扫描）")
        if value:
            self.msg_paths = list(dict.fromkeys([*self.msg_paths, value]))
            self.msg_path_var.set("；".join(self.msg_paths))

    def _clear_msg_paths(self) -> None:
        self.msg_paths = []
        self.msg_path_var.set("")

    def _choose_output(self) -> None:
        value = filedialog.askdirectory(title="选择归档保存位置")
        if value:
            self.output_root.set(value)

    def _append_log(self, text: str) -> None:
        self.log.configure(state="normal"); self.log.insert(tk.END, text.rstrip() + "\n"); self.log.see(tk.END); self.log.configure(state="disabled")

    def _load_folders(self) -> None:
        if self.process:
            messagebox.showinfo("正在运行", "请等待当前操作结束。")
            return
        if self.source_mode.get() == "msg":
            self.status.set("MSG 来源无需读取邮箱；请选择文件/文件夹后直接预览。")
            return
        args = [worker_python(), "-m", "outlook_archiver", "folders", "--source", self.source_mode.get(), "--output-json", str(self.folder_result)]
        if self.source_mode.get() == "pst":
            if not self.pst_path.get():
                messagebox.showerror("缺少 PST", "请先选择 PST 文件。")
                return
            if not self.pst_work_root.get().strip():
                self._update_default_pst_work_root()
            args += [
                "--pst", self.pst_path.get(),
                "--pst-work-root", self.pst_work_root.get().strip(),
                "--pst-reuse-root", self.output_root.get().strip(),
            ]
        self.status.set("正在读取文件夹；大型 PST 首次会制作并校验工作副本……")
        self._launch(args, "folders")

    def _show_folders(self) -> None:
        index = self.store_combo.current()
        if index < 0:
            return
        store = self.catalog["stores"][index]
        self.visible_folders = [f for f in self.catalog["folders"] if f["store_id"] == store["store_id"] and f.get("selectable", True)]
        self.folder_list.delete(0, tk.END)
        for folder in self.visible_folders:
            depth = max(0, len(folder.get("segments", [])) - 1)
            self.folder_list.insert(tk.END, "　" * depth + folder["folder_name"] + "　〔" + folder["folder_path"] + "〕")

    def _config(self) -> dict[str, Any]:
        store_index = self.store_combo.current()
        selected = list(self.folder_list.curselection())
        msg_mode = self.source_mode.get() == "msg"
        if msg_mode and not self.msg_paths:
            raise ValueError("请选择至少一个 MSG/MSD 文件或包含它们的文件夹。")
        if not msg_mode and (store_index < 0 or not selected):
            raise ValueError("请先读取邮箱，并选择至少一个邮件文件夹。")
        selected_store = {"store_id": "msg-import", "store_path": ""} if msg_mode else self.catalog["stores"][store_index]
        if not msg_mode and self.source_mode.get() == "outlook" and str(selected_store.get("store_path", "")).lower().endswith(".pst"):
            raise ValueError("当前选择是 PST 数据文件。请改选“PST 文件”来源，以使用经过校验的完整工作副本。")
        root = self.output_root.get().strip()
        if not root:
            raise ValueError("请选择保存位置。")
        config: dict[str, Any] = {
            "source_mode": self.source_mode.get(), "pst_path": self.pst_path.get().strip(),
            "pst_work_root": self.pst_work_root.get().strip(),
            "pst_work_root_auto": not self.pst_work_custom,
            "msg_paths": self.msg_paths,
            "store_id": selected_store["store_id"],
            "folder_ids": [] if msg_mode else [self.visible_folders[index]["folder_id"] for index in selected],
            "folder_paths": [] if msg_mode else [self.visible_folders[index]["folder_path"] for index in selected],
            "include_subfolders": self.include_subfolders.get(), "output_root": root,
            "layout": self.layout.get(), "include_addresses": self.include_addresses.get(), "years": [],
            "index_mode": self.index_mode.get(),
            "save_msg": self.save_msg.get(), "save_attachments": self.save_attachments.get(),
            "save_inline_images": self.save_inline_images.get(),
            "save_body_sources": any((self.save_msg.get(), self.save_attachments.get(), self.save_inline_images.get())),
            "test_limit": 5 if self.sample_limit.get() else None,
        }
        if self.date_mode.get() == "years":
            config["years"] = parse_years(self.years.get())
            if not config["years"]:
                raise ValueError("请输入至少一个年份。")
        elif self.date_mode.get() == "range":
            start = date.fromisoformat(self.start_date.get().strip()); end = date.fromisoformat(self.end_date.get().strip())
            if end < start:
                raise ValueError("结束日期不能早于开始日期。")
            config["start_date"] = start.isoformat()
            config["end_date"] = (end + timedelta(days=1)).isoformat()
        return config

    def _run_job(self, command: str) -> None:
        if self.process:
            messagebox.showinfo("正在运行", "请等待当前操作结束，或先点击停止。")
            return
        try:
            config = self._config()
            write_json_atomic(self.job_config, config)
        except Exception as exc:
            messagebox.showerror("设置不完整", str(exc)); return
        if self.stop_file.exists():
            self.stop_file.unlink()
        args = [worker_python(), "-m", "outlook_archiver", command, "--config", str(self.job_config)]
        if self.sample_limit.get() and command == "export":
            args += ["--limit", "5"]
        if command == "export":
            args += ["--stop-file", str(self.stop_file)]
        self.progress["value"] = 0; self.counter.set("成功 0　跳过 0　警告 0　失败 0")
        self.status.set("正在预览……" if command == "preview" else "正在导出……")
        self._launch(args, command)

    def _launch(self, args: list[str], purpose: str) -> None:
        self.pending_result = None
        self.completion_status.set({"export": "正在转换，请等待完成提示……",
                                    "preview": "正在预览……", "folders": "正在读取邮箱和文件夹……"}.get(purpose, "正在处理……"))
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        try:
            environment = dict(os.environ); environment["PYTHONUTF8"] = "1"
            self.process = subprocess.Popen(args, cwd=str(Path(__file__).resolve().parent.parent), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace", creationflags=flags, env=environment)
        except Exception as exc:
            self.process = None; messagebox.showerror("启动失败", str(exc)); return
        self.stop_button.configure(state="normal" if purpose == "export" else "disabled")
        threading.Thread(target=self._read_process, args=(self.process, purpose), daemon=True).start()

    def _read_process(self, process: subprocess.Popen[str], purpose: str) -> None:
        assert process.stdout
        for line in process.stdout:
            self.message_queue.put(("line", line.rstrip()))
        code = process.wait()
        self.message_queue.put(("done", (purpose, code)))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, value = self.message_queue.get_nowait()
                if kind == "line":
                    self._handle_line(value)
                else:
                    self._finished(*value)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _handle_line(self, line: str) -> None:
        if not line.startswith("EVENT "):
            self._append_log(line); return
        try:
            event = json.loads(line[6:])
        except json.JSONDecodeError:
            self._append_log(line); return
        name = event.get("event")
        if name == "progress":
            total = max(1, int(event.get("total", 1))); current = int(event.get("current", 0))
            self.progress["maximum"] = total; self.progress["value"] = current
            warning = int(event.get("warning", 0)) + int(event.get("partial", 0))
            failed = int(event.get("failed", 0)) + int(event.get("source_unavailable", 0))
            self.counter.set(f"成功 {event.get('success', 0)}　跳过 {event.get('skipped', 0)}　警告 {warning}　失败 {failed}")
            self.status.set(f"{current}/{total}　{event.get('folder', '')}")
        elif name == "scan_progress":
            self.status.set(f"正在扫描：{event.get('folder')}　{event.get('current')}/{event.get('total')}")
        elif name == "preview_complete":
            self.status.set(f"预览完成：选中 {event.get('selected', 0)} 封普通邮件；非邮件项目 {event.get('non_mail', 0)}；时间未知 {event.get('unknown_time', 0)}。")
            self._append_log(self.status.get())
            for warning in event.get("warnings", []): self._append_log("警告：" + warning)
        elif name == "run_complete":
            self.last_report = event.get("report", "")
            self.last_failures = event.get("failures", "")
            self.pending_result = event
            self.completion_status.set("结果报告已生成，正在结束任务……")
        elif name == "finalizing":
            self.completion_status.set("邮件处理结束，正在生成索引和结果报告，请稍候……")
        elif name == "report_ready":
            self.last_report = event["report"]
            self.last_failures = event["failures"]
            self.completion_status.set("结果报告已可查看；正在更新全归档索引，请勿关闭程序……")
            self._append_log("本次结果报告已生成：" + self.last_report)
        elif name == "index_progress":
            self.completion_status.set(f"报告已可查看；全归档索引已读取 {event['current']} 条，请稍候……")
        elif name == "index_mode":
            self.completion_status.set("报告已可查看；索引正在" + event['mode'] + "，请稍候……")
        elif name == "fatal":
            self.status.set("操作失败：" + event.get("error", "未知错误")); self._append_log(self.status.get())
            self.completion_status.set(self.status.get())
        else:
            self._append_log(json.dumps(event, ensure_ascii=False))

    def _finished(self, purpose: str, code: int) -> None:
        self.process = None; self.stop_button.configure(state="disabled")
        if purpose == "export" and code == 0:
            result = self.pending_result
            self.pending_result = None
            if result is None:
                self.status.set("任务已退出，但未收到完成结果，无法确认转换完成。请查看结果报告和窗口记录。")
                self.completion_status.set(self.status.get())
                messagebox.showwarning("未能确认完成", self.status.get(), parent=self)
                return
            title = completion_title(result)
            failed = sum(int(result.get(key, 0)) for key in ("failed", "partial", "source_unavailable"))
            summary = (f"{title}：成功 {result.get('success', 0)}，有警告 {result.get('warning', 0)}，"
                       f"跳过 {result.get('skipped', 0)}，失败/部分 {failed}，未处理 {result.get('unprocessed', 0)}。")
            self.status.set(summary)
            self.completion_status.set(summary)
            self.counter.set(summary)
            self._append_log(summary)
            self._show_result_dialog(result)
        elif purpose == "folders" and code == 0:
            try:
                self.catalog = read_json(self.folder_result)
                values = [item["store_name"] + (f"（{item['store_path']}）" if item.get("store_path") else "") for item in self.catalog["stores"]]
                self.store_combo["values"] = values
                if values:
                    self.store_combo.current(0); self._show_folders()
                note = "；PST 工作副本已复用" if self.catalog.get("pst_copy_reused") else ""
                self.status.set(f"已读取 {len(values)} 个邮箱/PST、{len(self.catalog['folders'])} 个文件夹{note}。")
                if self.catalog.get("pst_copy_reused") and self.catalog.get("pst_work_copy"):
                    self._append_log("已复用现有 PST 工作副本：" + str(self.catalog["pst_work_copy"]))
                for warning in self.catalog.get("warnings", []): self._append_log("警告：" + warning)
            except Exception as exc:
                self.status.set("文件夹结果读取失败：" + str(exc))
            self.completion_status.set(self.status.get())
        elif code != 0:
            self.pending_result = None
            self.completion_status.set("任务异常结束，未能确认完成。请查看窗口记录和结果报告。")
            messagebox.showerror("操作失败", self.status.get() + "\n\n详细信息请看窗口底部记录。", parent=self)
        else:
            self.completion_status.set(self.status.get())

    def _stop(self) -> None:
        if self.process:
            self.stop_file.write_text("stop\n", encoding="utf-8")
            self.status.set("已请求停止；当前单封完成后保存检查点并停止……")

    def _on_close(self) -> None:
        if self.process:
            if self.stop_button.instate(["!disabled"]):
                self._stop()
                messagebox.showinfo("正在安全停止", "已请求停止。请等待当前单封完成和结果报告写入后再关闭窗口。")
            else:
                messagebox.showinfo("操作仍在进行", "请等待当前读取/预览结束后再关闭窗口。")
            return
        self.destroy()

    def _open_output(self) -> None:
        path = Path(self.output_root.get()).expanduser()
        if path.exists(): os.startfile(path)
        else: messagebox.showinfo("尚无输出", "保存位置还不存在。")

    def _open_report(self) -> None:
        path = Path(self.last_report) if self.last_report else None
        if path and path.exists(): os.startfile(path)
        else:
            log_dir = Path(self.output_root.get()).expanduser() / "logs"
            reports = sorted(log_dir.glob("run_*_report.md"), reverse=True) if log_dir.exists() else []
            if reports: os.startfile(reports[0])
            else: messagebox.showinfo("尚无报告", "请先完成一次导出。")

    def _open_failures(self) -> None:
        path = Path(self.last_failures) if self.last_failures else None
        if path and path.exists():
            os.startfile(path)
            return
        log_dir = Path(self.output_root.get()).expanduser() / "logs"
        reports = sorted(log_dir.glob("run_*_failures.csv"), reverse=True) if log_dir.exists() else []
        if reports:
            os.startfile(reports[0])
        else:
            messagebox.showinfo("尚无失败清单", "请先完成一次导出。")

    def _show_result_dialog(self, event: dict[str, Any]) -> None:
        dialog = tk.Toplevel(self)
        dialog.title(completion_title(event))
        dialog.transient(self)
        dialog.resizable(False, False)
        body = ttk.Frame(dialog, padding=14)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=completion_title(event), font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", pady=(0, 10))
        failed = int(event.get("failed", 0)) + int(event.get("source_unavailable", 0))
        summary = (
            f"新完成：{event.get('success', 0)}\n"
            f"有警告：{event.get('warning', 0)}\n"
            f"已跳过：{event.get('skipped', 0)}\n"
            f"部分完成：{event.get('partial', 0)}\n"
            f"失败或源不可访问：{failed}\n"
            f"停止后未处理：{event.get('unprocessed', 0)}"
        )
        ttk.Label(body, text=summary, justify="left").pack(anchor="w")
        note = "失败清单包含主题、来源、文件夹、失败阶段、具体原因和重试提示。" if failed or int(event.get("partial", 0)) else "可打开结果报告核对本次处理范围和警告。"
        if event.get("scan_stopped") or int(event.get("unprocessed", 0)):
            note += " 本次尚未全部完成，可点击主窗口的‘继续’。"
        ttk.Label(body, text=note, wraplength=440).pack(anchor="w", pady=(10, 10))
        buttons = ttk.Frame(body)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="打开结果报告", command=self._open_report).pack(side="left")
        ttk.Button(buttons, text="打开失败清单", command=self._open_failures).pack(side="left", padx=6)
        ttk.Button(buttons, text="打开输出目录", command=self._open_output).pack(side="left")
        ttk.Button(buttons, text="关闭", command=dialog.destroy).pack(side="right")
        dialog.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - dialog.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - dialog.winfo_height()) // 2)
        dialog.geometry(f"+{x}+{y}")
        if self.state() == "iconic":
            self.deiconify()
        dialog.deiconify()
        dialog.lift()
        dialog.attributes("-topmost", True)
        dialog.after(1500, lambda: dialog.attributes("-topmost", False) if dialog.winfo_exists() else None)
        dialog.grab_set()
        dialog.focus_set()
        self.bell()


def run_gui() -> None:
    app = App()
    app.mainloop()
