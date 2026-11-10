"""Start the window, with a sweep when one is piped in. ip-table-database-editor.cmd runs this in console
python, which is the only kind that receives a pipe; it loads no Qt, so it is quick.

    pinger.exe -json 172.22.0.1/24 | ip-table-database-editor.cmd        the window, with the sweep loaded
    ip-table-database-editor.cmd [TABLE.xlsx | SWEEP.json] [--db ...]   no pipe: the window as before

With a pipe (or "-"), the sweep is read here, the window starts in pythonw, detached, and gets it on its own
stdin, and this returns, so the prompt is free at once. A table or sweep named on the command line wins over
a pipe.
"""
import os
import subprocess
import sys


def stdin_piped(stream=None):
    """True when stdin is a pipe or a redirected file, not a terminal or nothing. Git Bash's terminal (mintty)
    reaches a Windows program as a named pipe called like \\msys-…-pty0-from-master, so on Windows a pipe with
    such a name counts as a terminal; reading it would wait for typing."""
    stream = sys.stdin if stream is None else stream
    if stream is None:
        return False
    try:
        fd = stream.fileno()
    except (OSError, ValueError):
        return False
    if os.name != "nt":
        import stat
        mode = os.fstat(fd).st_mode
        return stat.S_ISFIFO(mode) or stat.S_ISREG(mode)      # not a terminal, not /dev/null
    import ctypes
    import msvcrt
    from ctypes import wintypes
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.GetFileType.argtypes = [wintypes.HANDLE]
    k32.GetFileInformationByHandleEx.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD]
    try:
        handle = msvcrt.get_osfhandle(fd)
    except OSError:
        return False
    kind = k32.GetFileType(handle)
    if kind == 1:                       # FILE_TYPE_DISK: < sweep.json
        return True
    if kind != 3:                       # FILE_TYPE_CHAR (a console, NUL) or unknown
        return False
    buf = ctypes.create_string_buffer(4 + 2 * 1024)
    if not k32.GetFileInformationByHandleEx(handle, 2, buf, len(buf)):       # FileNameInfo
        return True                     # an anonymous pipe (cmd's and PowerShell's |) can still say nothing
    length = int.from_bytes(buf.raw[:4], "little")
    name = buf.raw[4:4 + length].decode("utf-16-le", "replace")
    return not ("-pty" in name and ("msys-" in name or "cygwin-" in name))


def split_args(argv):
    """(the table or sweep named, the other arguments)."""
    named, rest, i = None, [], 0
    while i < len(argv):
        a = argv[i]
        if a == "--db" and i + 1 < len(argv):
            rest += argv[i:i + 2]
            i += 2
            continue
        if a.startswith("--db="):
            rest.append(a)
        elif named is None and (a == "-" or not a.startswith("-")):
            named = a
        else:
            rest.append(a)
        i += 1
    return named, rest


def pythonw():
    exe = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    return exe if os.path.exists(exe) else sys.executable


def preflight():
    """What would stop the window from starting, said here: the window runs detached in pythonw, with nowhere
    to print an error. Returns one line per problem, or an empty list."""
    import importlib.util
    from .oui import DEFAULT as OUI
    problems = ["%s is not installed for %s (pip install -r requirements.txt)" % (lib, sys.executable)
                for lib in ("PySide6", "openpyxl") if importlib.util.find_spec(lib) is None]
    if not os.path.exists(OUI):
        problems.append("the manufacturer list %s is missing: put pinger's oui.txt beside pinger, "
                        "or name it with IPTABLE_OUI" % OUI)
    return problems


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    problems = preflight()
    if problems:
        print("ip-table-database-editor can't start:\n  " + "\n  ".join(problems), file=sys.stderr)
        return 1
    named, rest = split_args(argv)
    data = None
    if named == "-" or (named is None and stdin_piped()):
        data = sys.stdin.buffer.read()
        named = "-"
    cmd = [pythonw(), "-m", "iptable.gui"] + ([named] if named else []) + (["--piped"] if data is not None else []) + rest
    flags = (subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP) if os.name == "nt" else 0
    child = subprocess.Popen(cmd, stdin=subprocess.PIPE if data is not None else subprocess.DEVNULL,
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=flags,
                             start_new_session=os.name != "nt")
    if data is not None:
        child.stdin.write(data)
        child.stdin.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
