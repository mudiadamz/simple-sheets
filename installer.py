"""Simple Sheets' setup / uninstall program: simplecore.installer with this app's name, file types and icon.
Silent: SimpleSheets-Setup.exe /S [/D=folder] [/desktop] [/noopenwith] [/relaunch]; uninstall: --uninstall."""
import version  # noqa: F401  (tells simplecore who the app is - before the installer reads it)

if __name__ == "__main__":
    from simplecore import installer
    installer.main()
