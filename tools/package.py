"""Build step after PyInstaller: dist/SimpleSheets-portable.zip (the app + an empty SimpleSheetsData folder, so
settings stay next to the exe) and dist/SimpleSheets-Setup.exe.sha256 (checked by automatic updates)."""
import hashlib
import os
import zipfile

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIST = os.path.join(HERE, "dist")


def main():
    with zipfile.ZipFile(os.path.join(DIST, "SimpleSheets-portable.zip"), "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(DIST, "SimpleSheets.exe"), "SimpleSheets/SimpleSheets.exe")
        z.writestr("SimpleSheets/SimpleSheetsData/README.txt",
                   "Settings, recent files and recovery copies are kept in this folder (portable mode).\n")
    setup = os.path.join(DIST, "SimpleSheets-Setup.exe")
    h = hashlib.sha256()
    with open(setup, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    with open(setup + ".sha256", "w", encoding="ascii") as f:
        f.write(f"{h.hexdigest()}  SimpleSheets-Setup.exe\n")
    print("Packaged: SimpleSheets-portable.zip, SimpleSheets-Setup.exe.sha256")


if __name__ == "__main__":
    main()
