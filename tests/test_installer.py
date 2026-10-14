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
        assert read(f"Classes\\Applications\\SimpleSheets.exe\\SupportedTypes", ext) == ""
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
