import io
import tempfile
import unittest
import wave
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from fsm_gui import App, detect_audio_format, get_audio_duration


class FakeVariable:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeTree:
    def __init__(self, columns, displaycolumns=("#all",)):
        self.columns = tuple(columns)
        self.displaycolumns = tuple(displaycolumns)
        self.config = {}
        self.selected = None
        self.focused = None

    def __getitem__(self, key):
        if key == "columns":
            return self.columns
        if key == "displaycolumns":
            return self.displaycolumns
        raise KeyError(key)

    def configure(self, **kwargs):
        self.config.update(kwargs)
        if "displaycolumns" in kwargs:
            self.displaycolumns = tuple(kwargs["displaycolumns"])

    def identify(self, region, _x, _y):
        return "cell" if region == "region" else ""

    def identify_column(self, _x):
        return "#1"

    def identify_row(self, _y):
        return "I001"

    def index(self, _item_id):
        return 0

    def selection_set(self, item_id):
        self.selected = item_id

    def focus(self, item_id):
        self.focused = item_id

    def get_children(self):
        return ("I001",)


class AudioFormatTests(unittest.TestCase):
    def test_detects_supported_audio_signatures(self):
        self.assertEqual(detect_audio_format(b"RIFF\x00\x00\x00\x00WAVE"), ("WAV", ".wav"))
        self.assertEqual(
            detect_audio_format(b"\x00\x00\x00\x1cftypM4A "),
            ("M4A", ".m4a"),
        )
        self.assertEqual(detect_audio_format(b"ID3\x04\x00"), ("MP3", ".mp3"))
        self.assertEqual(detect_audio_format(b"\xff\xfb\x90\x04"), ("MP3", ".mp3"))
        self.assertEqual(detect_audio_format(b"ikke lyd"), ("Ukjent", ""))

    def test_reads_wav_duration(self):
        data = io.BytesIO()
        with wave.open(data, "wb") as audio:
            audio.setnchannels(1)
            audio.setsampwidth(2)
            audio.setframerate(8000)
            audio.writeframes(b"\x00\x00" * 8000)

        duration = get_audio_duration(data.getvalue(), "WAV")

        self.assertAlmostEqual(duration, 1.0, places=3)

    def test_all_supplied_music_files_have_duration(self):
        music_zip = (
            Path(__file__).resolve().parent
            / "Musikk_02-09-2026_KUNSTLØP Oppvisning Bergen.zip"
        )
        if not music_zip.exists():
            self.skipTest("Musikk-ZIP er ikke tilgjengelig.")

        missing = []
        with zipfile.ZipFile(music_zip, "r") as archive:
            for entry in archive.infolist():
                if entry.is_dir():
                    continue
                data = archive.read(entry)
                audio_format, _suffix = detect_audio_format(data)
                if not get_audio_duration(data, audio_format):
                    missing.append(entry.filename)

        self.assertEqual(missing, [])


class AudioCacheTests(unittest.TestCase):
    def make_app(self, temp_dir, data):
        music_zip = Path(temp_dir) / "Musikk.zip"
        with zipfile.ZipFile(music_zip, "w") as archive:
            archive.writestr("mappe/sang.mp3", data)

        app = App.__new__(App)
        app.music_zip = music_zip
        app.music_cache_dir = Path(temp_dir) / "cache"
        app.music_cache_dir.mkdir()
        app.music_file_info = {
            "mappe/sang.mp3": {"format": "MP3", "suffix": ".mp3"}
        }
        return app

    def test_reuses_existing_cache_file_without_rewriting(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            app = self.make_app(temp_dir, b"ID3-lyddata")
            first_path = app.get_cached_mp3_path("mappe/sang.mp3")

            with patch.object(Path, "write_bytes") as write_bytes:
                second_path = app.get_cached_mp3_path("mappe/sang.mp3")

            self.assertEqual(second_path, first_path)
            write_bytes.assert_not_called()

    def test_changed_audio_content_gets_new_cache_path(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            app = self.make_app(temp_dir, b"ID3-forste-versjon")
            first_path = app.get_cached_mp3_path("mappe/sang.mp3")
            with zipfile.ZipFile(app.music_zip, "w") as archive:
                archive.writestr("mappe/sang.mp3", b"ID3-andre-versjon")

            second_path = app.get_cached_mp3_path("mappe/sang.mp3")

            self.assertNotEqual(second_path, first_path)
            self.assertEqual(second_path.read_bytes(), b"ID3-andre-versjon")


class TableInteractionTests(unittest.TestCase):
    def test_resolves_column_after_columns_are_hidden(self):
        app = App.__new__(App)
        app.tree = FakeTree(
            ("startnummer", "musikknavn"),
            displaycolumns=("musikknavn",),
        )

        column_name = app.get_tree_column_name(10)

        self.assertEqual(column_name, "musikknavn")

    def test_updates_visible_columns(self):
        app = App.__new__(App)
        app.table_columns = ("startnummer", "musikknavn")
        app.table_column_vars = {
            "startnummer": FakeVariable(False),
            "musikknavn": FakeVariable(True),
        }
        app.tree = FakeTree(app.table_columns)

        app.update_visible_table_columns("startnummer")

        self.assertEqual(app.tree.displaycolumns, ("musikknavn",))

    def test_prevents_hiding_every_column(self):
        app = App.__new__(App)
        app.table_columns = ("startnummer",)
        app.table_column_vars = {"startnummer": FakeVariable(False)}
        app.tree = FakeTree(app.table_columns)

        with patch("fsm_gui.messagebox.showinfo") as showinfo:
            app.update_visible_table_columns("startnummer")

        self.assertTrue(app.table_column_vars["startnummer"].get())
        showinfo.assert_called_once()

    def test_checkbox_requires_confirmation_and_marks_row_completed(self):
        app = App.__new__(App)
        app.tree = FakeTree(("lop_gjennomfort",))
        app.rows = [{"NavnFraIsonen": "Testløper"}]
        app.root = object()
        app.log = Mock()
        app.refresh_table = Mock()
        event = SimpleNamespace(x=10, y=10)

        with patch("fsm_gui.messagebox.askyesno", return_value=True) as askyesno:
            result = app.on_tree_checkbox_click(event)

        self.assertEqual(result, "break")
        self.assertTrue(app.rows[0]["LopGjennomfort"])
        askyesno.assert_called_once()
        app.refresh_table.assert_called_once()

    def test_checkbox_removal_can_be_cancelled(self):
        app = App.__new__(App)
        app.tree = FakeTree(("lop_gjennomfort",))
        app.rows = [{"NavnFraIsonen": "Testløper", "LopGjennomfort": True}]
        app.root = object()
        app.log = Mock()
        app.refresh_table = Mock()
        event = SimpleNamespace(x=10, y=10)

        with patch("fsm_gui.messagebox.askyesno", return_value=False) as askyesno:
            result = app.on_tree_checkbox_click(event)

        self.assertEqual(result, "break")
        self.assertTrue(app.rows[0]["LopGjennomfort"])
        askyesno.assert_called_once()
        app.refresh_table.assert_not_called()


if __name__ == "__main__":
    unittest.main()
