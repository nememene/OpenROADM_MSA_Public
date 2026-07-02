#!/usr/bin/env python3
"""GUI launcher: prompt for a folder path and export size report to Excel."""

from __future__ import annotations

import os
import subprocess
import sys
import threading
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:
    print(
        "错误：未安装 tkinter，无法打开图形界面。\n"
        "Windows 请重新安装 Python 并勾选 tcl/tk；"
        "Linux 请执行: sudo apt install python3-tk",
        file=sys.stderr,
    )
    raise SystemExit(1)

from folder_size_monitor import run_scan


def default_output_path(scan_root: Path) -> Path:
    safe_name = scan_root.name or "root"
    return scan_root / f"{safe_name}_folder_sizes.xlsx"


def choose_folder_dialog(parent: tk.Misc | None = None, title: str = "选择要扫描的文件夹") -> str | None:
    """Open a folder picker. Uses macOS native dialog on Darwin for reliability."""
    if sys.platform == "darwin":
        escaped_title = title.replace('"', '\\"')
        script = f'POSIX path of (choose folder with prompt "{escaped_title}")'
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=600,
            )
            if result.returncode == 0:
                path = result.stdout.strip()
                return path or None
        except (OSError, subprocess.SubprocessError):
            pass

    if parent is not None:
        parent.update_idletasks()
        parent.lift()
        parent.focus_force()

    selected = filedialog.askdirectory(
        parent=parent,
        title=title,
        initialdir=str(Path.home()),
        mustexist=True,
    )
    return selected or None


def choose_save_dialog(
    parent: tk.Misc | None = None,
    title: str = "选择 Excel 输出位置",
    default_name: str = "folder_sizes.xlsx",
) -> str | None:
    """Open a save-file picker. Uses macOS native dialog on Darwin for reliability."""
    if sys.platform == "darwin":
        escaped_title = title.replace('"', '\\"')
        escaped_name = default_name.replace('"', '\\"')
        script = (
            f'set savePath to choose file name with prompt "{escaped_title}" '
            f'default name "{escaped_name}"\n'
            "return POSIX path of savePath"
        )
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True,
                text=True,
                timeout=600,
            )
            if result.returncode == 0:
                path = result.stdout.strip()
                if path and not path.lower().endswith(".xlsx"):
                    path += ".xlsx"
                return path or None
        except (OSError, subprocess.SubprocessError):
            pass

    if parent is not None:
        parent.update_idletasks()
        parent.lift()
        parent.focus_force()

    selected = filedialog.asksaveasfilename(
        parent=parent,
        title=title,
        defaultextension=".xlsx",
        initialdir=str(Path.home()),
        initialfile=default_name,
        filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
    )
    return selected or None


class FolderSizeMonitorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("文件夹大小监控")
        self.root.resizable(False, False)
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=16)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frame, text="请输入要扫描的文件夹地址：").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )

        self.path_var = tk.StringVar()
        path_entry = ttk.Entry(frame, textvariable=self.path_var, width=52)
        path_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        path_entry.focus_set()

        ttk.Button(frame, text="浏览...", command=self._browse_folder).grid(
            row=1, column=2, sticky="e"
        )

        ttk.Label(frame, text="Excel 输出路径：").grid(row=2, column=0, sticky="w", pady=(12, 4))
        self.output_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.output_var, width=52).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=(0, 8)
        )
        ttk.Button(frame, text="选择...", command=self._browse_output).grid(
            row=3, column=2, sticky="e"
        )

        self.status_var = tk.StringVar(value="准备就绪")
        ttk.Label(frame, textvariable=self.status_var).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(12, 8)
        )

        self.progress = ttk.Progressbar(frame, mode="indeterminate", length=420)
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 12))

        button_row = ttk.Frame(frame)
        button_row.grid(row=6, column=0, columnspan=3, sticky="e")
        self.start_button = ttk.Button(button_row, text="开始扫描", command=self._start_scan)
        self.start_button.pack(side="right")
        ttk.Button(button_row, text="取消", command=self.root.destroy).pack(side="right", padx=(0, 8))

        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

    def _browse_folder(self) -> None:
        self.status_var.set("请选择文件夹...")
        self.root.update_idletasks()
        try:
            selected = choose_folder_dialog(self.root, title="选择要扫描的文件夹")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("错误", f"无法打开文件夹选择器：\n{exc}")
            self.status_var.set("准备就绪")
            return

        self.status_var.set("准备就绪")
        if selected:
            self.path_var.set(selected)
            if not self.output_var.get().strip():
                self.output_var.set(str(default_output_path(Path(selected))))

    def _browse_output(self) -> None:
        default_name = "folder_sizes.xlsx"
        current = self.output_var.get().strip()
        if current:
            default_name = Path(current).name

        self.status_var.set("请选择输出位置...")
        self.root.update_idletasks()
        try:
            selected = choose_save_dialog(
                self.root,
                title="选择 Excel 输出位置",
                default_name=default_name,
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("错误", f"无法打开文件保存选择器：\n{exc}")
            self.status_var.set("准备就绪")
            return

        self.status_var.set("准备就绪")
        if selected:
            self.output_var.set(selected)

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.start_button.configure(state=state)
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    def _start_scan(self) -> None:
        folder_text = self.path_var.get().strip()
        if not folder_text:
            messagebox.showwarning("提示", "请先输入或选择文件夹地址。")
            return

        scan_root = Path(folder_text).expanduser()
        output_text = self.output_var.get().strip()
        output_path = Path(output_text).expanduser() if output_text else default_output_path(scan_root)

        if not scan_root.exists():
            messagebox.showerror("错误", f"路径不存在：\n{scan_root}")
            return
        if not scan_root.is_dir():
            messagebox.showerror("错误", f"路径不是文件夹：\n{scan_root}")
            return

        self._set_busy(True)
        self.status_var.set(f"正在扫描：{scan_root}")

        def worker() -> None:
            try:
                folder_count, elapsed = run_scan(scan_root, output_path)
            except Exception as exc:  # noqa: BLE001 - show any scan failure to the user
                self.root.after(0, lambda: self._on_scan_failed(str(exc)))
                return
            self.root.after(
                0,
                lambda: self._on_scan_finished(folder_count, elapsed, output_path),
            )

        threading.Thread(target=worker, daemon=True).start()

    def _on_scan_failed(self, message: str) -> None:
        self._set_busy(False)
        self.status_var.set("扫描失败")
        messagebox.showerror("扫描失败", message)

    def _on_scan_finished(self, folder_count: int, elapsed: float, output_path: Path) -> None:
        self._set_busy(False)
        self.status_var.set("扫描完成")
        resolved_output = output_path.resolve()
        messagebox.showinfo(
            "扫描完成",
            f"共扫描 {folder_count} 个文件夹\n"
            f"耗时 {elapsed:.1f} 秒\n"
            f"结果已保存到：\n{resolved_output}",
        )
        self._open_output_folder(resolved_output)

    def _open_output_folder(self, output_path: Path) -> None:
        folder = output_path.parent
        try:
            if sys.platform == "win32":
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(folder)], check=False)
            else:
                subprocess.run(["xdg-open", str(folder)], check=False)
        except OSError:
            pass


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    root.update_idletasks()
    root.deiconify()
    if sys.platform == "darwin":
        root.createcommand("tk::mac::ReopenApplication", root.deiconify)
    FolderSizeMonitorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
