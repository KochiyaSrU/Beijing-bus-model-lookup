"""打包后冒烟测试：启动 EXE，等主窗口出现，可选截图，用退出码报告结果。

用法（在程序目录下执行）：:

    python verify_exe.py                     # 正常启动 + 离线启动，两次都验
    python verify_exe.py 别的名字.exe         # 检查指定文件
    python verify_exe.py --shot out.png      # 顺便把主窗口截一张图
    python verify_exe.py --timeout 30        # 改等待上限（秒）
    python verify_exe.py --offline           # 只跑离线那一次

退出码：0 = 窗口正常出现，1 = 超时没出现，2 = 找不到文件。
"""

from __future__ import annotations

import ctypes
import os
import subprocess
import sys
import tempfile
import time
from ctypes import wintypes
from pathlib import Path

APP_TITLE = "北京公交车型查询器"
DEFAULT_EXE = f"{APP_TITLE}.exe"

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
EnumWindows = user32.EnumWindows
EnumWindows.argtypes = [ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM), wintypes.LPARAM]
GetWindowTextW = user32.GetWindowTextW
GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
GetWindowRect = user32.GetWindowRect
GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]


def find_window(title: str = APP_TITLE) -> int:
    """返回标题匹配的顶层窗口句柄（没有则 0）。"""
    found: list[int] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def visit(window, _lparam):
        buffer = ctypes.create_unicode_buffer(512)
        GetWindowTextW(window, buffer, 512)
        if buffer.value == title:
            found.append(window)
            return False
        return True

    EnumWindows(visit, 0)
    return found[0] if found else 0


def screenshot(window: int, target: Path) -> bool:
    """把窗口内容存成 PNG（需要 Pillow）。"""
    try:
        from PIL import ImageGrab
    except ImportError:
        print("  (没有安装 Pillow，跳过截图)")
        return False
    user32.SetForegroundWindow(window)
    time.sleep(1.0)
    rect = wintypes.RECT()
    if not GetWindowRect(window, ctypes.byref(rect)):
        return False
    image = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
    target.parent.mkdir(parents=True, exist_ok=True)
    image.save(target)
    print(f"  截图已保存：{target}（{image.width}x{image.height}）")
    return True


def launch_and_wait(target_exe: Path, workdir: Path, timeout: int, *, offline: bool, shot_path: Path | None) -> bool:
    """启动一次 EXE，等主窗口出现；返回是否成功。offline=True 时用 BPT_NO_NETWORK=1 强制离线。"""
    temp_dir = Path(tempfile.mkdtemp(prefix="bpt-verify-"))
    environment = dict(os.environ)
    environment["TMP"] = environment["TEMP"] = str(temp_dir)
    if offline:
        environment["BPT_NO_NETWORK"] = "1"
    label = "离线（BPT_NO_NETWORK=1）" if offline else "正常"
    print(f"启动 {target_exe.name}（{label}）…")
    process = subprocess.Popen([str(target_exe)], cwd=str(workdir), env=environment)
    started = time.time()
    window = 0
    try:
        for _ in range(timeout):
            time.sleep(1)
            if process.poll() is not None:
                print(f"程序在 {time.time() - started:.1f} 秒后自行退出（退出码 {process.returncode}）。")
                return False
            window = find_window()
            if window:
                print(f"主窗口在 {time.time() - started:.1f} 秒后出现。")
                break
        else:
            print(f"{timeout} 秒内没有等到主窗口。")
            return False
        if shot_path is not None:
            screenshot(window, shot_path)
        return True
    finally:
        subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"], capture_output=True)
        time.sleep(1)
        try:
            temp_dir.rmdir()
        except OSError:
            pass


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    target_exe = Path(DEFAULT_EXE)
    shot_path: Path | None = None
    timeout = 45
    offline_only = False
    index = 0
    while index < len(args):
        if args[index] == "--shot" and index + 1 < len(args):
            shot_path = Path(args[index + 1])
            index += 2
        elif args[index] == "--timeout" and index + 1 < len(args):
            timeout = int(args[index + 1])
            index += 2
        elif args[index] == "--offline":
            offline_only = True
            index += 1
        else:
            target_exe = Path(args[index])
            index += 1

    if not target_exe.is_file():
        print(f"找不到 {target_exe}，请在程序目录下运行，或把 EXE 路径作为参数传进来。")
        return 2

    workdir = Path(__file__).resolve().parent
    # 联网功能接进来之后，"离线也能正常用"是硬要求，所以默认两次都跑：
    # 正常启动 ✓ + BPT_NO_NETWORK=1 启动 ✓（--offline 只跑离线那一次）
    checks = [(True, None)] if offline_only else [(False, shot_path), (True, None)]
    for offline, shot in checks:
        if not launch_and_wait(target_exe, workdir, timeout, offline=offline, shot_path=shot):
            return 1
    print("冒烟测试通过。" + ("" if offline_only else "（正常 + 离线两次都通过）"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
