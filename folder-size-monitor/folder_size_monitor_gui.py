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
    from tkinter import filedialog, messagebox
except ImportError:
    print(
        "错误：未安装 tkinter，无法打开图形界面。\n"
        "Windows 请重新安装 Python 并勾选 tcl/tk；"
        "Linux 请执行: sudo apt install python3-tk",
        file=sys.stderr,
    )
    raise SystemExit(1)

from folder_size_monitor import run_scan


def configure_macos_tk(root: tk.Tk) -> None:
    """Fix Retina display scaling so buttons receive clicks correctly."""
    if sys.platform != "darwin":
        return

    os.environ.setdefault("TK_SILENCE_DEPRECATION", "1")
    try:
        scaling = root.winfo_fpixels("1i") / 72.0
        root.tk.call("tk", "scaling", scaling)
    except tk.TclError:
        pass


def show_alert(title: str, message: str, alert_type: str = "info") -> None:
    """Show an alert dialog. Uses macOS native dialog on Darwin for reliability."""
    if sys.platform == "darwin":
        escaped_title = title.replace('"', '\\"')
        escaped_message = message.replace('"', '\\"').replace("\n", "\\n")
        icon = {"info": "note", "warning": "caution", "error": "stop"}.get(alert_type, "note")
        script = (
            f'display dialog "{escaped_message}" '
            f'with title "{escaped_title}" '
            f'buttons {{"OK"}} default button "OK" '
            f'with icon {icon}'
        )
        try:
            subprocess.run(["osascript", "-e", script], check=False)
            return
        except OSError:
            pass

    if alert_type == "warning":
        messagebox.showwarning(title, message)
    elif alert_type == "error":
        messagebox.showerror(title, message)
    else:
        messagebox.showinfo(title, message)


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
        self.root.configure(bg="#2b2b2b")
        configure_macos_tk(self.root)
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.root.bind("<Return>", lambda _event: self._start_scan())
        self.root.bind("<Escape>", lambda _event: self._on_cancel())

    def _make_button(self, parent: tk.Misc, text: str, command, primary: bool = False) -> tk.Button:
        bg = "#0a84ff" if primary else "#555555"
        active_bg = "#0066cc" if primary else "#666666"
        button = tk.Button(
            parent,
            text=text,
            command=command,
            width=12,
            height=1,
            padx=12,
            pady=8,
            bg=bg,
            fg="white",
            activebackground=active_bg,
            activeforeground="white",
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            cursor="hand2",
            font=("PingFang SC", "Helvetica", "Arial", 13),
        )
        return button

    def _build_ui(self) -> None:
        frame = tk.Frame(self.root, padx=20, pady=20, bg="#2b2b2b")
        frame.pack(fill=tk.BOTH, expand=True)

        label_font = ("PingFang SC", "Helvetica", "Arial", 13)
        entry_font = ("Menlo", "Monaco", "Courier", 12)

        tk.Label(
            frame,
            text="请输入要扫描的文件夹地址：",
            bg="#2b2b2b",
            fg="#f0f0f0",
            font=label_font,
            anchor="w",
        ).pack(fill=tk.X, pady=(0, 8))

        path_row = tk.Frame(frame, bg="#2b2b2b")
        path_row.pack(fill=tk.X, pady=(0, 12))

        self.path_var = tk.StringVar()
        path_entry = tk.Entry(
            path_row,
            textvariable=self.path_var,
            width=48,
            font=entry_font,
            bg="#1e1e1e",
            fg="#ffffff",
            insertbackground="#ffffff",
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#555555",
            highlightcolor="#0a84ff",
        )
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        path_entry.focus_set()

        self._make_button(path_row, "浏览...", self._browse_folder).pack(side=tk.RIGHT)

        tk.Label(
            frame,
            text="Excel 输出路径：",
            bg="#2b2b2b",
            fg="#f0f0f0",
            font=label_font,
            anchor="w",
        ).pack(fill=tk.X, pady=(0, 8))

        output_row = tk.Frame(frame, bg="#2b2b2b")
        output_row.pack(fill=tk.X, pady=(0, 12))

        self.output_var = tk.StringVar()
        tk.Entry(
            output_row,
            textvariable=self.output_var,
            width=48,
            font=entry_font,
            bg="#1e1e1e",
            fg="#ffffff",
            insertbackground="#ffffff",
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground="#555555",
            highlightcolor="#0a84ff",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self._make_button(output_row, "选择...", self._browse_output).pack(side=tk.RIGHT)

        self.status_var = tk.StringVar(value="准备就绪")
        tk.Label(
            frame,
            textvariable=self.status_var,
            bg="#2b2b2b",
            fg="#aaaaaa",
            font=label_font,
            anchor="w",
        ).pack(fill=tk.X, pady=(0, 8))

        self.progress = tk.Canvas(frame, height=6, bg="#1e1e1e", highlightthickness=0)
        self.progress.pack(fill=tk.X, pady=(0, 16))
        self._progress_bar = self.progress.create_rectangle(0, 0, 0, 6, fill="#0a84ff", width=0)
        self._progress_animating = False

        button_row = tk.Frame(frame, bg="#2b2b2b")
        button_row.pack(fill=tk.X)

        self.start_button = self._make_button(button_row, "开始扫描", self._start_scan, primary=True)
        self.start_button.pack(side=tk.RIGHT)

        self.cancel_button = self._make_button(button_row, "取消", self._on_cancel)
        self.cancel_button.pack(side=tk.RIGHT, padx=(0, 10))

        self.root.update_idletasks()
        self.root.minsize(self.root.winfo_width(), self.root.winfo_height())

    def _on_cancel(self) -> None:
        self.status_var.set("正在退出...")
        self.root.update_idletasks()
        self.root.quit()
        self.root.destroy()

    def _browse_folder(self) -> None:
        self.status_var.set("请选择文件夹...")
        self.root.update_idletasks()
        try:
            selected = choose_folder_dialog(self.root, title="选择要扫描的文件夹")
        except Exception as exc:  # noqa: BLE001
            self.status_var.set("打开文件夹选择器失败")
            show_alert("错误", f"无法打开文件夹选择器：\n{exc}", "error")
            return

        if selected:
            self.path_var.set(selected)
            if not self.output_var.get().strip():
                self.output_var.set(str(default_output_path(Path(selected))))
            self.status_var.set(f"已选择：{selected}")
        else:
            self.status_var.set("准备就绪")

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
            self.status_var.set("打开保存选择器失败")
            show_alert("错误", f"无法打开文件保存选择器：\n{exc}", "error")
            return

        if selected:
            self.output_var.set(selected)
            self.status_var.set(f"输出到：{selected}")
        else:
            self.status_var.set("准备就绪")

    def _set_busy(self, busy: bool) -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        self.start_button.configure(state=state)
        self.cancel_button.configure(state=state)
        if busy:
            self._animate_progress()
        else:
            self._progress_animating = False
            self.progress.coords(self._progress_bar, 0, 0, 0, 6)

    def _animate_progress(self) -> None:
        self._progress_animating = True
        width = self.progress.winfo_width() or 420

        def step(position: int = 0) -> None:
            if not self._progress_animating:
                return
            bar_width = max(40, width // 5)
            start = position % (width + bar_width)
            self.progress.coords(self._progress_bar, start - bar_width, 0, start, 6)
            self.root.after(60, lambda: step(position + 20))

        step()

    def _start_scan(self) -> None:
        folder_text = self.path_var.get().strip()
        if not folder_text:
            self.status_var.set("请先输入或选择文件夹地址")
            show_alert("提示", "请先输入或选择文件夹地址。", "warning")
            return

        scan_root = Path(folder_text).expanduser()
        output_text = self.output_var.get().strip()
        output_path = Path(output_text).expanduser() if output_text else default_output_path(scan_root)

        if not scan_root.exists():
            self.status_var.set("路径不存在")
            show_alert("错误", f"路径不存在：\n{scan_root}", "error")
            return
        if not scan_root.is_dir():
            self.status_var.set("路径不是文件夹")
            show_alert("错误", f"路径不是文件夹：\n{scan_root}", "error")
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
        show_alert("扫描失败", message, "error")

    def _on_scan_finished(self, folder_count: int, elapsed: float, output_path: Path) -> None:
        self._set_busy(False)
        resolved_output = output_path.resolve()
        self.status_var.set(f"扫描完成：{resolved_output}")
        show_alert(
            "扫描完成",
            f"共扫描 {folder_count} 个文件夹\n"
            f"耗时 {elapsed:.1f} 秒\n"
            f"结果已保存到：\n{resolved_output}",
            "info",
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
    if sys.platform == "darwin":
        root.createcommand("tk::mac::ReopenApplication", lambda: root.deiconify())
    FolderSizeMonitorApp(root)
    root.lift()
    root.attributes("-topmost", True)
    root.after(200, lambda: root.attributes("-topmost", False))
    root.focus_force()
    root.mainloop()


if __name__ == "__main__":
    main()
