from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from src.html_splitter import HtmlSplitConfig, split_html
from src.sheet_copier import list_sheet_names
from src.splitter import SplitConfig, split_workbook

MODE_XLSX = "XLSX"
MODE_HTML = "HTML"

SUBTITLE_XLSX = (
    "XLSX를 HEADER + 1행 단위로 분할합니다. "
    "결과는 INPUT과 같은 경로의 {원본파일명}_{timestamp} 폴더에 저장됩니다."
)
SUBTITLE_HTML = (
    "HTML 테이블을 헤더 + 1행 단위로 분할합니다. "
    "결과는 INPUT과 같은 경로의 {원본파일명}_{timestamp} 폴더에 저장됩니다."
)


class MemoSplitterApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("MEMO SPLITTER")
        self.geometry("780x620")
        self.minsize(700, 520)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.mode_var = tk.StringVar(value=MODE_XLSX)
        self.input_var = tk.StringVar()
        self.multimedia_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.header_row_var = tk.IntVar(value=1)
        self.progress_var = tk.DoubleVar(value=0.0)
        self.status_var = tk.StringVar(value="대기 중")
        self._worker: threading.Thread | None = None

        self._build_layout()
        self._apply_mode()

    def _build_layout(self) -> None:
        title = ctk.CTkLabel(self, text="MEMO SPLITTER", font=ctk.CTkFont(size=22, weight="bold"))
        title.pack(anchor="w", padx=16, pady=(16, 4))

        mode_frame = ctk.CTkFrame(self, fg_color="transparent")
        mode_frame.pack(anchor="w", padx=16, pady=(0, 4))
        ctk.CTkLabel(mode_frame, text="모드").pack(side="left", padx=(0, 8))
        self.mode_segment = ctk.CTkSegmentedButton(
            mode_frame,
            values=[MODE_XLSX, MODE_HTML],
            variable=self.mode_var,
            command=self._on_mode_change,
        )
        self.mode_segment.pack(side="left")

        self.subtitle = ctk.CTkLabel(
            self,
            text=SUBTITLE_XLSX,
            font=ctk.CTkFont(size=13),
        )
        self.subtitle.pack(anchor="w", padx=16, pady=(0, 12))

        self.form = ctk.CTkFrame(self)
        self.form.pack(fill="x", padx=16, pady=8)
        self.form.grid_columnconfigure(1, weight=1)

        self.input_label = ctk.CTkLabel(self.form, text="INPUT XLSX")
        self.input_label.grid(row=0, column=0, sticky="w", padx=12, pady=10)
        self.input_entry = ctk.CTkEntry(self.form, textvariable=self.input_var)
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=12, pady=10)
        self.input_button = ctk.CTkButton(
            self.form, text="찾기", width=70, command=self._browse_input
        )
        self.input_button.grid(row=0, column=2, padx=12, pady=10)

        self.multimedia_label = ctk.CTkLabel(self.form, text="Multimedia 폴더")
        self.multimedia_label.grid(row=1, column=0, sticky="w", padx=12, pady=10)
        self.multimedia_entry = ctk.CTkEntry(self.form, textvariable=self.multimedia_var)
        self.multimedia_entry.grid(row=1, column=1, sticky="ew", padx=12, pady=10)
        self.multimedia_button = ctk.CTkButton(
            self.form, text="폴더", width=70, command=self._browse_multimedia
        )
        self.multimedia_button.grid(row=1, column=2, padx=12, pady=10)

        self.sheet_label = ctk.CTkLabel(self.form, text="SHEET")
        self.sheet_label.grid(row=2, column=0, sticky="w", padx=12, pady=10)
        self.sheet_combo = ctk.CTkComboBox(self.form, variable=self.sheet_var, values=[""])
        self.sheet_combo.grid(row=2, column=1, sticky="ew", padx=12, pady=10)

        self.header_label = ctk.CTkLabel(self.form, text="HEADER ROW")
        self.header_label.grid(row=3, column=0, sticky="w", padx=12, pady=10)
        self.header_spin = ctk.CTkEntry(self.form, textvariable=self.header_row_var, width=80)
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

    def _on_mode_change(self, _value: str | None = None) -> None:
        self.input_var.set("")
        self.sheet_var.set("")
        self.sheet_combo.configure(values=[""])
        self._apply_mode()

    def _apply_mode(self) -> None:
        is_xlsx = self.mode_var.get() == MODE_XLSX
        if is_xlsx:
            self.subtitle.configure(text=SUBTITLE_XLSX)
            self.input_label.configure(text="INPUT XLSX")
            self.multimedia_label.grid()
            self.multimedia_entry.grid()
            self.multimedia_button.grid()
            self.sheet_label.grid()
            self.sheet_combo.grid()
            self.header_label.grid()
            self.header_spin.grid()
        else:
            self.subtitle.configure(text=SUBTITLE_HTML)
            self.input_label.configure(text="INPUT HTML")
            self.multimedia_label.grid_remove()
            self.multimedia_entry.grid_remove()
            self.multimedia_button.grid_remove()
            self.sheet_label.grid_remove()
            self.sheet_combo.grid_remove()
            self.header_label.grid_remove()
            self.header_spin.grid_remove()

    def _browse_input(self) -> None:
        if self.mode_var.get() == MODE_HTML:
            path = filedialog.askopenfilename(
                title="INPUT HTML 선택",
                filetypes=[("HTML files", "*.html;*.htm"), ("All files", "*.*")],
            )
            if not path:
                return
            self.input_var.set(path)
            return

        path = filedialog.askopenfilename(
            title="INPUT XLSX 선택",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        self.input_var.set(path)
        self._load_sheet_names(Path(path))

    def _browse_multimedia(self) -> None:
        path = filedialog.askdirectory(title="Multimedia 폴더 선택")
        if path:
            self.multimedia_var.set(path)

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

        input_path = Path(self.input_var.get().strip())

        if self.mode_var.get() == MODE_HTML:
            config: SplitConfig | HtmlSplitConfig = HtmlSplitConfig(
                input_path=input_path,
            )
            worker_target = self._run_html_split_thread
        else:
            try:
                header_row = int(self.header_row_var.get())
            except (tk.TclError, ValueError):
                messagebox.showerror("입력 오류", "HEADER ROW는 숫자여야 합니다.")
                return

            sheet_name = self.sheet_var.get().strip()
            if not sheet_name:
                messagebox.showerror("입력 오류", "SHEET를 선택하세요.")
                return

            config = SplitConfig(
                input_path=input_path,
                multimedia_root=Path(self.multimedia_var.get().strip()),
                sheet_name=sheet_name,
                header_row=header_row,
            )
            worker_target = self._run_xlsx_split_thread

        self.run_button.configure(state="disabled")
        self.progress_var.set(0)
        self.status_var.set("SPLIT 실행 중...")
        self._clear_log()

        self._worker = threading.Thread(
            target=worker_target,
            args=(config,),
            daemon=True,
        )
        self._worker.start()

    def _run_xlsx_split_thread(self, config: SplitConfig) -> None:
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

        self.after(0, self._on_split_complete, result, MODE_XLSX)

    def _run_html_split_thread(self, config: HtmlSplitConfig) -> None:
        try:
            result = split_html(
                config,
                on_log=lambda message: self.after(0, self._append_log, message),
                on_progress=lambda current, total: self.after(
                    0, self._update_progress, current, total
                ),
            )
        except Exception as exc:
            self.after(0, self._on_split_failed, str(exc))
            return

        self.after(0, self._on_split_complete, result, MODE_HTML)

    def _update_progress(self, current: int, total: int) -> None:
        if total <= 0:
            self.progress_var.set(0)
            return
        self.progress_var.set(current / total)

    def _on_split_complete(self, result, mode: str = MODE_XLSX) -> None:
        self.run_button.configure(state="normal")
        self.progress_var.set(1.0)
        output_info = f", 출력 {result.output_root}" if result.output_root else ""
        parts = [
            f"완료: {'파일' if mode == MODE_HTML else '폴더'} {result.folders_created}개",
            f"첨부 {result.attachments_copied}개",
        ]
        if mode == MODE_XLSX:
            parts.append(f"이미지 임베드 {result.images_embedded}개")
        if result.attachment_skips:
            parts.append(f"첨부 스킵 {len(result.attachment_skips)}건")
        summary = ", ".join(parts) + output_info
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
