"""控制台编码：把 stdout / stderr 切到 UTF-8。

Windows 下输出被管道接走时，Python 默认按本地代码页（如 cp936）编码，
Git Bash、IDE 任务输出里就看到一串乱码。Windows 控制台本身走 Unicode API，
重设成 UTF-8 不影响交互式显示，只是让管道里的字节也是 UTF-8。
"""
import sys


def setup_console():
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is None:
                continue
            enc = (stream.encoding or "").lower().replace("_", "-")
            if enc not in ("utf-8", "utf8"):
                stream.reconfigure(encoding="utf-8")
        except (AttributeError, OSError, ValueError):
            # 被重定向到不支持 reconfigure 的流时保持原样，不要为编码问题中断任务。
            pass
