from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from src.attachment_copier import list_sheet_names
from src.splitter import SplitConfig, split_workbook


class MemoSplitterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("MEMO SPLITTER")
        self.geometry("760x560")
        self.minsize(680, 480)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.header_row_var = tk.IntVar(value=1)
        self.progress_var = tk.DoubleVar(value=0.0)
        self.status_var = tk.StringVar(value="대기 중")
        self._worker: threading.Thread | None = None

        self._build_layout()

    def _build_layout(self) -> None:
        padding = {"padx": 16, "pady": 8}

        title = ctk.CTkLabel(self, text="MEMO SPLITTER", font=ctk.CTkFont(size=22, weight="bold"))
        title.pack(anchor="w", padx=16, pady=(16, 4))

        subtitle = ctk.CTkLabel(
            self,
            text="XLSX를 HEADER + 1행 단위로 분할하고, 행별 첨부파일을 함께 복사합니다.",
            font=ctk.CTkFont(size=13),
        )
        subtitle.pack(anchor="w", padx=16, pady=(0, 12))

        form = ctk.CTkFrame(self)
        form.pack(fill="x", padx=16, pady=8)
        form.grid_columnconfigure(1, weight=1)

        self._add_path_row(form, 0, "INPUT XLSX", self.input_var, self._browse_input, is_file=True)
        self._add_path_row(form, 1, "OUTPUT 폴더", self.output_var, self._browse_output, is_file=False)

        ctk.CTkLabel(form, text="SHEET").grid(row=2, column=0, sticky="w", padx=12, pady=10)
        self.sheet_combo = ctk.CTkComboBox(form, variable=self.sheet_var, values=[""])
        self.sheet_combo.grid(row=2, column=1, sticky="ew", padx=12, pady=10)

        ctk.CTkLabel(form, text="HEADER ROW").grid(row=3, column=0, sticky="w", padx=12, pady=10)
        self.header_spin = ctk.CTkEntry(form, textvariable=self.header_row_var, width=80)
        self.header_spin.grid(row=3, column=1, sticky="w", padx=12, pady=10)

        self.run_button = ctk.CTkButton(self, text="SPLIT 실행", command=self._start_split)
        self.run_button.pack(padx=16, pady=8, anchor="w")

        self.progress = ctk.CTkProgressBar(self, variable=self.progress_var)
        self.progress.pack(fill="x", padx=16, pady=(4, 4))
        self.progress.set(0)

        self.status_label = ctk.CTkLabel(self, textvariable=self.status_var, anchor="w")
        self.status_label.pack(fill="x", padx=16, pady=(0, 8))

        log_frame = ctk.CTkFrame(self)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        self.log_box = ctk.CTkTextbox(log_frame, wrap="word")
        self.log_box.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        self.log_box.configure(state="disabled")

    def _add_path_row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
        is_file: bool,
    ) -> None:
        ctk.CTkLabel(parent, text=label).grid(row=row, column=0, sticky="w", padx=12, pady=10)
        entry = ctk.CTkEntry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=12, pady=10)
        button_label = "찾기" if is_file else "폴더"
        ctk.CTkButton(parent, text=button_label, width=70, command=command).grid(
            row=row, column=2, padx=12, pady=10
        )

    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="INPUT XLSX 선택",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        self.input_var.set(path)
        self._load_sheet_names(Path(path))

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="OUTPUT 폴더 선택")
        if path:
            self.output_var.set(path)

    def _load_sheet_names(self, xlsx_path: Path) -> None:
        try:
            sheets = list_sheet_names(xlsx_path)
        except Exception as exc:
            messagebox.showerror("시트 로드 실패", str(exc))
            return

        self.sheet_combo.configure(values=sheets)
        if sheets:
            self.sheet_var.set(sheets[0])
        self._append_log(f"시트 목록 로드: {', '.join(sheets)}")

    def _start_split(self) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showwarning("실행 중", "이미 SPLIT 작업이 진행 중입니다.")
            return

        try:
            header_row = int(self.header_row_var.get())
        except (tk.TclError, ValueError):
            messagebox.showerror("입력 오류", "HEADER ROW는 숫자여야 합니다.")
            return

        config = SplitConfig(
            input_path=Path(self.input_var.get().strip()),
            output_root=Path(self.output_var.get().strip()),
            sheet_name=self.sheet_var.get().strip(),
            header_row=header_row,
        )

        if not config.sheet_name:
            messagebox.showerror("입력 오류", "SHEET를 선택하세요.")
            return

        self.run_button.configure(state="disabled")
        self.progress_var.set(0)
        self.status_var.set("SPLIT 실행 중...")
        self._clear_log()

        self._worker = threading.Thread(
            target=self._run_split_thread,
            args=(config,),
            daemon=True,
        )
        self._worker.start()

    def _run_split_thread(self, config: SplitConfig) -> None:
        try:
            result = split_workbook(
                config,
                on_log=lambda message: self.after(0, self._append_log, message),
                on_progress=lambda current, total: self.after(
                    0, self._update_progress, current, total
                ),
            )
        except Exception as exc:
            self.after(0, self._on_split_failed, str(exc))
            return

        self.after(0, self._on_split_complete, result)

    def _update_progress(self, current: int, total: int) -> None:
        if total <= 0:
            self.progress_var.set(0)
            return
        self.progress_var.set(current / total)

    def _on_split_complete(self, result) -> None:
        self.run_button.configure(state="normal")
        self.progress_var.set(1.0)
        summary = (
            f"완료: 폴더 {result.folders_created}개, "
            f"첨부 {result.attachments_copied}개, "
            f"빈 행 스킵 {result.rows_skipped}개, "
            f"첨부 스킵 {len(result.attachment_skips)}건"
        )
        self.status_var.set(summary)
        self._append_log(summary)
        if result.row_errors:
            self._append_log("행 오류:")
            for error in result.row_errors:
                self._append_log(f"  - {error}")
        messagebox.showinfo("SPLIT 완료", summary)

    def _on_split_failed(self, message: str) -> None:
        self.run_button.configure(state="normal")
        self.status_var.set("실패")
        self._append_log(f"오류: {message}")
        messagebox.showerror("SPLIT 실패", message)

    def _clear_log(self) -> None:
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _append_log(self, message: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")


def run_app() -> None:
    app = MemoSplitterApp()
    app.mainloop()
