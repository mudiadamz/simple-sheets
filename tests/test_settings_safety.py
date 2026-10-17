"""Tests must never write the user's real settings (that once wiped a saved AI key on every test run)."""
import os

from simplecore import recent, settings


def real_home():
    return os.path.normcase(os.path.join(os.environ.get("APPDATA", ""), "SimpleSheets"))


def test_settings_folder_is_not_the_real_one():
    assert os.path.normcase(recent.config_dir()) != real_home()
    assert not os.path.normcase(recent.config_dir()).startswith(real_home())


def test_writing_settings_in_a_test_stays_in_the_test(tmp_path):
    settings.set("ai_key", "plain:from-a-test")
    path = os.path.join(recent.config_dir(), "settings.json")
    assert os.path.exists(path) and "from-a-test" in open(path, encoding="utf-8").read()
    real = os.path.join(real_home(), "settings.json")
    if os.path.exists(real):
        assert "from-a-test" not in open(real, encoding="utf-8").read()


def test_without_the_app_fixture_too(app_free=None):
    # this test has no `app` fixture on purpose: the autouse fixture alone must keep it away from %APPDATA%
    assert os.path.normcase(recent.config_dir()) != real_home()
