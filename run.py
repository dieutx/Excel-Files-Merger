from __future__ import annotations

from collections.abc import Callable
from datetime import datetime
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from excel_merger import (
    MERGE_MODE_COMBINE,
    MERGE_MODE_SEPARATE,
    MergeProgressEvent,
    discover_excel_files,
    merge_excel_files,
)


class ExcelMergerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel Files Merger")
        self.root.minsize(760, 520)

        self.source_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.merge_mode = tk.StringVar(value=MERGE_MODE_COMBINE)
        self.summary = tk.StringVar(value="Choose a source folder to scan for Excel workbooks.")
        self.progress_percent = tk.StringVar(value="0%")
        self._is_running = False
        self._widgets_disabled_during_merge: list[tk.Widget] = []

        self._configure_style()
        self._build_layout()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("Title.TLabel", font=("TkDefaultFont", 17, "bold"))
        style.configure("Section.TLabelframe.Label", font=("TkDefaultFont", 10, "bold"))
        style.configure("Summary.TLabel", padding=(10, 8))
        style.configure("Primary.TButton", padding=(14, 8))

    def _build_layout(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=16)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(4, weight=1)

        ttk.Label(main, text="Excel Files Merger", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )

        paths = ttk.LabelFrame(main, text="Files", style="Section.TLabelframe", padding=12)
        paths.grid(row=1, column=0, sticky="ew", pady=(14, 10))
        paths.columnconfigure(1, weight=1)

        source_entry, source_button = self._path_row(
            paths,
            0,
            "Source folder",
            self.source_path,
            "Browse...",
            self.choose_source_folder,
        )
        output_entry, output_button = self._path_row(
            paths,
            1,
            "Output file",
            self.output_path,
            "Save as...",
            self.choose_output_file,
        )
        self._widgets_disabled_during_merge.extend(
            [source_entry, source_button, output_entry, output_button]
        )

        mode_frame = ttk.LabelFrame(
            main, text="Merge Mode", style="Section.TLabelframe", padding=12
        )
        mode_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        mode_frame.columnconfigure(0, weight=1)

        combine_radio = ttk.Radiobutton(
            mode_frame,
            text="Combine all files into one worksheet.",
            variable=self.merge_mode,
            value=MERGE_MODE_COMBINE,
            command=self.refresh_summary,
        )
        combine_radio.grid(row=0, column=0, sticky="w")

        separate_radio = ttk.Radiobutton(
            mode_frame,
            text="Keep source sheets separate.",
            variable=self.merge_mode,
            value=MERGE_MODE_SEPARATE,
            command=self.refresh_summary,
        )
        separate_radio.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self._widgets_disabled_during_merge.extend([combine_radio, separate_radio])

        ttk.Label(main, textvariable=self.summary, style="Summary.TLabel", anchor="w").grid(
            row=3, column=0, sticky="ew", pady=(0, 10)
        )

        log_frame = ttk.LabelFrame(main, text="Log", style="Section.TLabelframe", padding=8)
        log_frame.grid(row=4, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame,
            height=9,
            wrap="word",
            state="disabled",
            borderwidth=0,
            padx=8,
            pady=8,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        actions = ttk.Frame(main)
        actions.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        actions.columnconfigure(0, weight=1)

        progress_frame = ttk.Frame(actions)
        progress_frame.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        progress_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            progress_frame,
            mode="determinate",
            maximum=100,
            value=0,
        )
        self.progress.grid(row=0, column=0, sticky="ew")

        ttk.Label(progress_frame, textvariable=self.progress_percent, width=5, anchor="e").grid(
            row=0, column=1, sticky="e", padx=(8, 0)
        )

        self.merge_button = ttk.Button(
            actions,
            text="Merge",
            style="Primary.TButton",
            command=self.merge,
        )
        self.merge_button.grid(row=0, column=1, sticky="e")

    def _path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        button_text: str,
        command: Callable[[], None],
    ) -> tuple[ttk.Entry, ttk.Button]:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10))

        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 8), pady=4)
        entry.bind("<FocusOut>", lambda _event: self.refresh_summary())
        entry.bind("<Return>", lambda _event: self.refresh_summary())

        button = ttk.Button(parent, text=button_text, command=command)
        button.grid(row=row, column=2, sticky="e", pady=4)
        return entry, button

    def choose_source_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select source folder")
        if not folder:
            return

        self.source_path.set(folder)
        self._append_log(f"Source folder: {folder}")
        self.refresh_summary()

    def choose_output_file(self) -> None:
        source = self.source_path.get().strip()
        initial_dir = source if source else str(Path.home())

        output = filedialog.asksaveasfilename(
            title="Save merged workbook",
            initialdir=initial_dir,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if not output:
            return

        self.output_path.set(output)
        self._append_log(f"Output file: {output}")
        self.refresh_summary()

    def refresh_summary(self) -> None:
        source = self.source_path.get().strip()
        output = self.output_path.get().strip()
        mode = self.merge_mode.get()

        if not source:
            self.summary.set("Choose a source folder to scan for Excel workbooks.")
            return

        try:
            files = discover_excel_files(source, exclude_path=output or None)
        except Exception as exc:
            self.summary.set(f"Unable to scan the source folder: {exc}")
            return

        mode_text = (
            "combine all files into one worksheet"
            if mode == MERGE_MODE_COMBINE
            else "keep source sheets separate"
        )
        output_text = output if output else "choose an output workbook before merging"
        self.summary.set(
            f"Found {len(files)} Excel workbook(s). Mode: {mode_text}. Output: {output_text}"
        )

    def merge(self) -> None:
        if self._is_running:
            return

        source = self.source_path.get().strip()
        output = self.output_path.get().strip()

        if not source:
            messagebox.showerror(
                "Source folder required",
                "Please choose the folder that contains the Excel files to merge.",
            )
            return
        if not output:
            messagebox.showerror(
                "Output file required",
                "Please choose where the merged Excel workbook should be saved.",
            )
            return

        self.refresh_summary()
        self._set_running(True)
        self._append_log("Starting merge...")

        worker = threading.Thread(
            name="excel-merge-worker",
            target=self._merge_worker,
            args=(source, output, self.merge_mode.get()),
        )
        worker.start()

    def _merge_worker(self, source: str, output: str, mode: str) -> None:
        try:
            result = merge_excel_files(
                source,
                output,
                mode,
                on_progress=lambda event: self.root.after(0, self._handle_progress_event, event),
            )
        except Exception as exc:
            self.root.after(0, self._merge_failed, str(exc))
            return

        self.root.after(0, self._merge_finished, result)

    def _handle_progress_event(self, event: MergeProgressEvent) -> None:
        if event.total_files > 0 and event.event_type in {"file_start", "sheet_processed", "file_complete"}:
            completed = event.current_file_index - 1
            if event.event_type == "file_complete":
                completed = event.current_file_index
            percent = int((completed / event.total_files) * 100)
        elif event.event_type == "complete":
            percent = 100
        else:
            percent = int(self.progress["value"])

        self.progress.configure(value=percent)
        self.progress_percent.set(f"{percent}%")

        status = self._format_progress_status(event)
        if status:
            self.summary.set(status)

        log_message = self._format_progress_log_message(event)
        if log_message:
            self._append_log(log_message)

    def _format_progress_status(self, event: MergeProgressEvent) -> str:
        if event.event_type == "scan_complete":
            return f"Found {event.total_files} Excel workbook(s). Preparing merge..."
        if event.event_type == "file_start" and event.file_path:
            return f"Processing {event.current_file_index}/{event.total_files}: {event.file_path.name}"
        if event.event_type == "sheet_processed" and event.file_path and event.sheet_title:
            return (
                f"Processed sheet '{event.sheet_title}' from "
                f"{event.file_path.name} ({event.current_file_index}/{event.total_files})."
            )
        if event.event_type == "writing_output":
            return "Writing merged workbook..."
        if event.event_type == "complete":
            return "Merge complete."
        return self.summary.get()

    def _format_progress_log_message(self, event: MergeProgressEvent) -> str:
        if event.event_type == "scan_complete":
            return f"Found {event.total_files} Excel workbook(s)."
        if event.event_type == "file_start" and event.file_path:
            return f"Processing file {event.current_file_index}/{event.total_files}: {event.file_path.name}"
        if event.event_type == "sheet_processed" and event.file_path and event.sheet_title:
            return (
                f"Processed sheet '{event.sheet_title}' from {event.file_path.name}; "
                f"{event.sheets_processed} sheet(s), {event.rows_written} row(s) so far."
            )
        if event.event_type == "file_complete" and event.file_path:
            return f"Finished file {event.current_file_index}/{event.total_files}: {event.file_path.name}"
        if event.event_type == "writing_output":
            return "Writing output workbook..."
        return ""

    def _merge_finished(self, result: object) -> None:
        self._set_running(False)
        self.refresh_summary()

        output = getattr(result, "output_path")
        message = (
            "Merge complete: "
            f"{getattr(result, 'files_processed')} file(s), "
            f"{getattr(result, 'sheets_processed')} sheet(s), "
            f"{getattr(result, 'rows_written')} row(s).\n"
            f"Output: {output}"
        )
        self.summary.set(message.replace("\n", " "))
        self._append_log(message)
        messagebox.showinfo("Merge complete", message)

    def _merge_failed(self, error: str) -> None:
        self._set_running(False)
        self.summary.set("Merge failed. See the log for details.")
        self._append_log(f"Merge failed: {error}")
        messagebox.showerror(
            "Merge failed",
            f"The selected files could not be merged.\n\n{error}",
        )

    def _set_running(self, is_running: bool) -> None:
        self._is_running = is_running
        state = "disabled" if is_running else "normal"
        self.merge_button.configure(state=state)
        for widget in self._widgets_disabled_during_merge:
            widget.configure(state=state)

        if is_running:
            self.progress.configure(value=0)
            self.progress_percent.set("0%")
            self.summary.set("Preparing merge...")
        else:
            if int(float(self.progress["value"])) < 100:
                self.progress_percent.set(f"{int(float(self.progress['value']))}%")

    def _on_close(self) -> None:
        if self._is_running:
            messagebox.showwarning(
                "Merge in progress",
                "Please wait for the current merge to finish before closing the app.",
            )
            return

        self.root.destroy()

    def _append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


def main() -> None:
    root = tk.Tk()
    app = ExcelMergerApp(root)
    app.refresh_summary()
    root.mainloop()


if __name__ == "__main__":
    main()
