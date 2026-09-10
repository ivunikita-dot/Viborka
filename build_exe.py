import os
import sys

def find_tcltk():
    try:
        import tkinter
        # try near tkinter package
        base = os.path.normpath(os.path.join(os.path.dirname(tkinter.__file__), '..', '..', 'tcl'))
        tcl8 = os.path.join(base, 'tcl8.6')
        tk8 = os.path.join(base, 'tk8.6')
        if os.path.isdir(tcl8) and os.path.isdir(tk8):
            return [(tcl8, 'tcl\\tcl8.6'), (tk8, 'tk\\tk8.6')]
        # fallback to sys.prefix
        p2 = os.path.join(sys.prefix, 'tcl')
        tcl8b = os.path.join(p2, 'tcl8.6')
        tk8b = os.path.join(p2, 'tk8.6')
        if os.path.isdir(tcl8b) and os.path.isdir(tk8b):
            return [(tcl8b, 'tcl\\tcl8.6'), (tk8b, 'tk\\tk8.6')]
    except Exception:
        return []
    return []


def main():
    script = 'find_fio_gui_recursive_v5.py'
    if not os.path.exists(script):
        print('Error: source script not found:', script, file=sys.stderr)
        sys.exit(1)

    args = ['--noconfirm', '--onefile', '--windowed', '--name', 'выборка', script]

    tcltk = find_tcltk()
    for src, dest in reversed(tcltk):
        args.insert(0, src + ';' + dest)
        args.insert(0, '--add-data')

    if os.path.exists('doom.ico'):
        args.insert(0, 'doom.ico')
        args.insert(0, '--icon')

    print('PyInstaller args:', args)

    try:
        import PyInstaller.__main__
    except Exception as e:
        print('PyInstaller not installed in this environment:', e, file=sys.stderr)
        sys.exit(1)

    PyInstaller.__main__.run(args)

if __name__ == '__main__':
    main()
