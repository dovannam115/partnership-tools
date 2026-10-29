# One-click install, run by "Install.bat" from the extracted zip.
# Copies the tool to %LOCALAPPDATA%\Programs\Partnership-Tools, then adds "Partnership Tools" to the Start menu
# (so Windows search finds it and it can be pinned to the taskbar) and to the Desktop.
# Running it again over an existing install replaces the code only: output\, app\dump
# and app\update.json stay as they are.
import os, sys, stat, time, shutil, socket, subprocess

SRC  = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEST = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser(r"~\AppData\Local"),
                    "Programs", "Partnership-Tools")
NAME = "Partnership Tools"
PORT = 8713
NO_WINDOW = 0x08000000

KEEP_TOP = {"output", "dist", ".git", "__pycache__"}      # never copied over
KEEP_APP = {"dump", "__pycache__"}


def running():
    s = socket.socket(); s.settimeout(0.5)
    try:
        s.connect(("127.0.0.1", PORT)); return True
    except Exception:
        return False
    finally:
        s.close()


def ignore(folder, names):
    rel = os.path.relpath(folder, SRC)
    out = set()
    for n in names:
        if rel == "." and n in KEEP_TOP: out.add(n)
        elif rel == "app" and (n in KEEP_APP or n.startswith("_backup_")): out.add(n)
        elif n == "__pycache__": out.add(n)
    if rel == "app" and os.path.exists(os.path.join(DEST, "app", "update.json")): out.add("update.json")
    return out


def put(src, dst):
    # An old copy can be read-only (zip/OneDrive attribute) or held for a moment by antivirus:
    # clear the flag, retry a few times, and never carry the flag over to the new copy.
    for i in range(5):
        try:
            if os.path.exists(dst): os.chmod(dst, stat.S_IWRITE)
            shutil.copy2(src, dst)
            os.chmod(dst, stat.S_IWRITE)
            return dst
        except PermissionError:
            if i == 4: raise
            time.sleep(1)


def copy():
    os.makedirs(DEST, exist_ok=True)
    shutil.copytree(SRC, DEST, ignore=ignore, copy_function=put, dirs_exist_ok=True)
    for d in ("output", os.path.join("app", "dump")):
        os.makedirs(os.path.join(DEST, d), exist_ok=True)


def shortcuts(dirs="@([Environment]::GetFolderPath('Programs'), [Environment]::GetFolderPath('Desktop'))"):
    target = os.path.join(DEST, "app", "pyembed", "pythonw.exe")
    args   = '"%s"' % os.path.join(DEST, "app", "server.py")
    icon   = os.path.join(DEST, "app", "ui", "app.ico")
    ps = ("$s = New-Object -ComObject WScript.Shell;"
          "foreach ($d in %s) {"
          " $l = $s.CreateShortcut((Join-Path $d '%s.lnk'));"
          " $l.TargetPath = '%s'; $l.Arguments = '%s'; $l.WorkingDirectory = '%s';"
          " $l.IconLocation = '%s,0'; $l.Description = 'Partnership Tools - PVI Digital'; $l.Save() }"
          % (dirs, NAME, target, args, DEST, icon))
    subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                   check=True, creationflags=NO_WINDOW)


def main():
    same = os.path.normcase(os.path.abspath(SRC)) == os.path.normcase(os.path.abspath(DEST))
    if running():
        print(NAME + " is open. Close it (Quit button in the app), then run Install.bat again.")
        return 1
    if not same:
        print("Copying to " + DEST + " (about 450 MB, takes a minute) ...")
        try:
            copy()
        except shutil.Error as ex:
            bad = [e[1] for e in ex.args[0]]
            print("Could not overwrite these files (open in another program, or blocked):")
            for b in bad[:10]: print("  " + b)
            if any(os.path.basename(b).lower() != "install.bat" for b in bad):
                print("Close Partnership Tools and any window showing that folder, then run Install.bat again.")
                return 1
            print("Only the installer copy was skipped, the app itself is in place.")
    print("Adding shortcuts (Start menu + Desktop) ...")
    shortcuts()
    subprocess.Popen([os.path.join(DEST, "app", "pyembed", "pythonw.exe"),
                      os.path.join(DEST, "app", "server.py")], cwd=DEST, creationflags=NO_WINDOW)
    print("Done. Open it any time: press Windows, type 'Partnership' and pick '%s'." % NAME)
    print("To pin: right-click '%s' in the Start menu > Pin to taskbar." % NAME)
    if not same:
        print("You can delete the extracted folder now: " + SRC)
    return 0


if __name__ == "__main__":
    sys.exit(main())
