"""Installer: files, "Open with" registration, uninstall. Uses a scratch registry key + temp folder,
never the real Classes / Uninstall keys."""
import os
import sys

import pytest

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Windows registry")

BASE = r"Software\SimpleSheetsTest"


@pytest.fixture
def inst(monkeypatch, tmp_path):
    import installer
    fake_exe = tmp_path / "payload.exe"
    fake_exe.write_bytes(b"MZ fake app")
    monkeypatch.setattr(installer, "payload", lambda: str(fake_exe))
    monkeypatch.setattr(installer, "_notify_explorer", lambda: None)
    yield installer
    installer._delete_tree("", BASE)            # remove the whole scratch key


def read(path, name=""):
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, f"{BASE}\\{path}") as k:
        return winreg.QueryValueEx(k, name)[0]


def exists(path):
    import winreg
    try:
        winreg.OpenKey(winreg.HKEY_CURRENT_USER, f"{BASE}\\{path}").Close()
        return True
    except FileNotFoundError:
        return False


def test_install_registers_open_with(inst, tmp_path):
    d = str(tmp_path / "app")
    inst.install(d, base=BASE, shortcuts=False)
    exe = os.path.join(d, "SimpleSheets.exe")
    assert os.path.exists(exe)
    for ext in inst.EXTS:
        assert read(f"Classes\\{ext}\\OpenWithProgids", inst.PROGID) == ""        # listed in "Open with"
        assert read("Classes\\Applications\\SimpleSheets.exe\\SupportedTypes", ext) == ""
    assert read(f"Classes\\{inst.PROGID}\\shell\\open\\command") == f'"{exe}" "%1"'
    assert read(r"Classes\Applications\SimpleSheets.exe", "FriendlyAppName") == "Simple Sheets"
    un = r"Microsoft\Windows\CurrentVersion\Uninstall\SimpleSheets"
    assert read(un, "DisplayName") == "Simple Sheets" and read(un, "DisplayVersion") == inst.APP_VERSION
    assert read(un, "UninstallString").endswith('Uninstall.exe" --uninstall')


def test_install_without_open_with(inst, tmp_path):
    inst.install(str(tmp_path / "app"), open_with=False, base=BASE, shortcuts=False)
    assert not exists(r"Classes\.xlsx\OpenWithProgids")
    assert exists(r"Microsoft\Windows\CurrentVersion\Uninstall\SimpleSheets")   # still uninstallable


def test_uninstall_removes_everything(inst, tmp_path):
    d = str(tmp_path / "app")
    inst.install(d, base=BASE, shortcuts=False)
    inst._set(r"Classes\.xlsx\OpenWithProgids", "Excel.Sheet.12", "", BASE)   # another app's entry
    inst.uninstall(d, base=BASE, shortcuts=False)
    assert not os.path.exists(d)
    assert not exists(f"Classes\\{inst.PROGID}") and not exists(r"Classes\Applications\SimpleSheets.exe")
    assert not exists(r"Microsoft\Windows\CurrentVersion\Uninstall\SimpleSheets")
    assert read(r"Classes\.xlsx\OpenWithProgids", "Excel.Sheet.12") == ""     # others untouched
    with pytest.raises(FileNotFoundError):
        read(r"Classes\.xlsx\OpenWithProgids", inst.PROGID)


def test_reinstall_over_existing(inst, tmp_path):
    d = str(tmp_path / "app")
    inst.install(d, base=BASE, shortcuts=False)
    inst.install(d, base=BASE, shortcuts=False)              # upgrade in place
    assert os.path.exists(os.path.join(d, "SimpleSheets.exe"))


def test_app_window_icon_is_bundled():
    import simple_sheets
    assert os.path.exists(simple_sheets.resource("assets", "simple_sheets.ico"))


def test_new_setup_removes_old_version(inst, tmp_path, monkeypatch):
    old_dir = str(tmp_path / "old place")
    inst.install(old_dir, base=BASE, shortcuts=False)
    with open(os.path.join(old_dir, "stale.txt"), "w") as f:      # a file only the old version had
        f.write("x")
    inst._set(r"Microsoft\Windows\CurrentVersion\Uninstall\SimpleSheets", "DisplayVersion", "0.9.0", BASE)
    assert inst.existing(BASE) == {"dir": old_dir, "version": "0.9.0"}
    new_dir = str(tmp_path / "new place")
    inst.install(new_dir, base=BASE, shortcuts=False)
    assert not os.path.exists(old_dir)                              # old version removed first
    assert os.path.exists(os.path.join(new_dir, "SimpleSheets.exe"))
    assert inst.existing(BASE) == {"dir": new_dir, "version": inst.APP_VERSION}
    inst.install(new_dir, base=BASE, shortcuts=False)              # same folder again: fine
    assert os.path.exists(os.path.join(new_dir, "SimpleSheets.exe"))


def test_running_old_version_blocks_install(inst, tmp_path, monkeypatch):
    d = str(tmp_path / "app")
    inst.install(d, base=BASE, shortcuts=False)
    monkeypatch.setattr(inst, "in_use", lambda folder: True)
    with pytest.raises(PermissionError):
        inst.install(str(tmp_path / "other"), base=BASE, shortcuts=False)
    assert os.path.exists(os.path.join(d, "SimpleSheets.exe"))     # nothing removed
