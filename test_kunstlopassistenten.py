import io
import tempfile
import unittest
import wave
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from datetime import datetime

from kunstlopassistenten import (
    App,
    add_missing_segments,
    auto_group_size,
    build_startliste,
    detect_audio_format,
    find_music_for_row,
    generate_startliste_excel,
    generate_vlc_playlist,
    get_audio_duration,
    group_overflows,
    layout_items,
    new_participant,
    parse_duration_mmss,
    parse_music_path,
    place_water_pauses,
    split_segments,
    validate_setup,
    water_pause_positions,
)


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


class ScheduleTests(unittest.TestCase):
    def runners(self, count):
        return [new_participant(f"Løper{i}", "Test", "Loddefjord IL") for i in range(1, count + 1)]

    def test_groups_warmup_and_times(self):
        items = layout_items(self.runners(3), 2, 2, 0)
        entries = build_startliste(
            items, 190, datetime(2000, 1, 1, 8, 15), ready_seconds=30, warmup_seconds=240
        )

        self.assertEqual(
            [(e["navn"], e["start"], e["end"], e["duration"]) for e in entries],
            [
                ("Oppvarmingsgruppe 1", "08:15:00", "08:19:00", "00:04:00"),
                ("Løper1 Test", "08:19:00", "08:22:40", "00:03:40"),
                ("Løper2 Test", "08:22:40", "08:26:20", "00:03:40"),
                ("Oppvarmingsgruppe 2", "08:26:20", "08:30:20", "00:04:00"),
                ("Løper3 Test", "08:30:20", "08:34:00", "00:03:40"),
            ],
        )
        self.assertEqual([e["nr"] for e in entries if not e["is_group"]], [1, 2, 3])

    def test_cancelled_runners_are_hidden_last(self):
        rows = self.runners(3)
        rows[0]["Påmelding"] = "Avmeldt"

        items = layout_items(rows, 1, 8, 0)

        self.assertIs(items[-1], rows[0])
        entries = build_startliste(items, 60, datetime(2000, 1, 1, 8, 0))
        self.assertEqual([e["navn"] for e in entries if e["kind"] == "run"], ["Løper2 Test", "Løper3 Test"])

    def test_water_pauses_between_groups(self):
        self.assertEqual(water_pause_positions(4, 1), [2])
        self.assertEqual(water_pause_positions(4, 2), [1, 3])
        self.assertEqual(water_pause_positions(1, 1), [])

        items = layout_items(self.runners(4), 2, 2, 1)
        entries = build_startliste(items, 60, datetime(2000, 1, 1, 8, 0), water_seconds=900)

        self.assertEqual([e["kind"] for e in entries],
                         ["warmup", "run", "run", "pause", "warmup", "run", "run"])
        self.assertEqual((entries[3]["start"], entries[3]["end"]), ("08:02:00", "08:17:00"))

    def test_no_warmup_groups(self):
        items = layout_items(self.runners(4), 0, 8, 1)

        self.assertEqual([i.get("Kind", "run") for i in items], ["run", "run", "pause", "run", "run"])

    def test_group_count_splits_runners(self):
        self.assertEqual(auto_group_size(7, 2), 4)
        self.assertEqual(auto_group_size(9, 3), 3)

        items = layout_items(self.runners(7), 3, 2, 0)
        self.assertEqual([i.get("Kind", "run") for i in items],
                         ["warmup", "run", "run", "warmup", "run", "run", "warmup", "run", "run", "run"])

        items = layout_items(self.runners(3), 1, 1, 0)
        self.assertEqual([i.get("Kind", "run") for i in items], ["warmup", "run", "run", "run"])

    def test_custom_run_length(self):
        rows = self.runners(2)
        rows[0]["LopLengde"] = 150
        entries = build_startliste(layout_items(rows, 0, 8, 0), 190, datetime(2000, 1, 1, 8, 0),
                                   ready_seconds=30)

        self.assertEqual([(e["run_length"], e["duration"], e["custom_length"]) for e in entries],
                         [("00:02:30", "00:03:00", True), ("00:03:10", "00:03:40", False)])

    def test_detects_group_with_too_many_runners(self):
        items = layout_items(self.runners(4), 2, 2, 0)
        self.assertEqual(group_overflows(items, 2), {})

        runner = items.pop(4)  # første utøver i gruppe 2 flyttes inn i gruppe 1
        items.insert(1, runner)

        self.assertEqual(group_overflows(items, 2), {id(items[0]): 3})

    def test_excel_has_original_sheet_and_sheet_with_duration(self):
        import openpyxl

        items = layout_items(self.runners(3), 2, 2, 0)
        entries = build_startliste(items, 190, datetime(2000, 1, 1, 8, 15), ready_seconds=30,
                                   warmup_seconds=240)
        with tempfile.TemporaryDirectory() as temp_dir:
            path = str(Path(temp_dir) / "startliste.xlsx")
            self.assertTrue(generate_startliste_excel(entries, path, "Tittel", Mock()))
            wb = openpyxl.load_workbook(path)

        self.assertEqual(wb.sheetnames, ["Startliste", "Med tidsforbruk"])
        plain, timed = wb["Startliste"], wb["Med tidsforbruk"]
        self.assertEqual([c.value for c in plain[3]],
                         ["Nr.", "Start", None, "Slutt", "Lengde på løp", "Navn", "Klubb"])
        self.assertEqual([c.value for c in plain[4]][1:4], ["08:15:00", "-", "08:19:00"])
        self.assertEqual(plain["F4"].value, "Oppvarmingsgruppe 1")
        self.assertEqual(plain["B7"].value, "ca. 08:26:20")
        self.assertEqual([c.value for c in timed[3]],
                         ["Nr.", "Start", None, "Slutt", "Lengde på løp", "Sammenlagt tidsforbruk",
                          "Navn", "Klubb"])
        self.assertEqual([c.value for c in timed[5]], [1, "08:19:00", "-", "08:22:40", "00:03:10",
                                                       "00:03:40", "Løper1 Test", "Loddefjord IL"])

    def test_moved_water_pause_keeps_its_place_until_count_changes(self):
        items = layout_items(self.runners(4), 2, 2, 1)
        pause = items.pop(3)
        items.insert(1, pause)

        entries = build_startliste(items, 60, datetime(2000, 1, 1, 8, 0), water_seconds=600)
        self.assertEqual(entries[1]["kind"], "pause")

        replaced = place_water_pauses(items, 1)
        self.assertEqual([i.get("Kind", "run") for i in replaced],
                         ["warmup", "run", "run", "pause", "warmup", "run", "run"])


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

        with patch("kunstlopassistenten.messagebox.showinfo") as showinfo:
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

        with patch("kunstlopassistenten.messagebox.askyesno", return_value=True) as askyesno:
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

        with patch("kunstlopassistenten.messagebox.askyesno", return_value=False) as askyesno:
            result = app.on_tree_checkbox_click(event)

        self.assertEqual(result, "break")
        self.assertTrue(app.rows[0]["LopGjennomfort"])
        askyesno.assert_called_once()
        app.refresh_table.assert_not_called()


class RegressionTests(unittest.TestCase):
    def runner(self, given, family="Test", **fields):
        row = new_participant(given, family)
        row.update(fields)
        return row

    def test_rejects_negative_and_invalid_durations(self):
        self.assertEqual(parse_duration_mmss("3:10"), 190)
        self.assertEqual(parse_duration_mmss("1:02:03"), 3723)
        for value in ("-1:00", "1:-05", "3:75", "1:60:00", "abc", "1"):
            self.assertIsNone(parse_duration_mmss(value), value)

    def test_stale_selection_after_list_is_replaced(self):
        app = App.__new__(App)
        app.schedule_tree = SimpleNamespace(selection=lambda: ("i5",))
        app.items = [self.runner("A")]

        self.assertIsNone(app.selected_item())

    def test_ambiguous_music_match_is_marked_uncertain(self):
        index = {f: parse_music_path(f) for f in (
            "Stevne/Junior_Jenter/FS/Anna_Berg_FS.mp3",
            "Stevne/Senior_Damer/FS/Anna_Berg_FS.mp3",
        )}
        row = self.runner("Anna", "Berg", Segment="FS", Klasse="Ukjent klasse")

        matched, uncertain = find_music_for_row(row, index, set())

        self.assertTrue(matched)
        self.assertTrue(uncertain)

    def test_music_match_uses_class_and_is_certain(self):
        index = {f: parse_music_path(f) for f in (
            "Stevne/Junior_Jenter/FS/Anna_Berg_FS.mp3",
            "Stevne/Senior_Damer/FS/Anna_Berg_FS.mp3",
        )}
        row = self.runner("Anna", "Berg", Segment="FS", Klasse="Senior")

        self.assertEqual(find_music_for_row(row, index, set()),
                         ("Stevne/Senior_Damer/FS/Anna_Berg_FS.mp3", False))

    def test_exact_names_are_matched_before_fuzzy_names(self):
        app = App.__new__(App)
        app.music_files = ["S/Klasse/FS/Anna_Berg_Hansen_FS.mp3"]
        app.music_index = {f: parse_music_path(f) for f in app.music_files}
        app.music_file_info, app.music_durations, app.music_status = {}, {}, {}
        fuzzy_row = self.runner("Anna", "Berg", Segment="FS")
        exact_row = self.runner("Anna", "Berg Hansen", Segment="FS")

        app.match_rows([fuzzy_row, exact_row], set())

        self.assertEqual(exact_row["MusikkFil"], app.music_files[0])
        self.assertEqual(fuzzy_row["MusikkFil"], "")

    def test_late_short_program_info_adds_free_skate(self):
        row = self.runner("Anna")
        items = [{"Kind": "warmup"}] + split_segments([row])
        row["Ovelse"] = "Kortprogram+Friløp"

        self.assertEqual(add_missing_segments(items), 1)
        self.assertEqual([(i.get("Segment"), i.get("Program")) for i in items[1:]],
                         [("SP", "Kortprogram"), ("FS", "Friløp")])
        self.assertEqual(add_missing_segments(items), 0)

    def test_sorting_keeps_cancelled_runners_out_of_active_slots(self):
        app = App.__new__(App)
        app.refresh_schedule = Mock()
        rows = [self.runner("C"), self.runner("A"), self.runner("B", Påmelding="Avmeldt")]
        app.items = layout_items(rows, 1, 8, 0)

        app.sort_by_given()

        self.assertEqual([r["GivenName"] for r in app.rows], ["A", "C", "B"])

    def test_validates_setup_before_loading(self):
        valid = {"type": "kunstlopassistenten-oppsett", "version": 1, "settings": {"start": "18:00"},
                 "items": [{"Kind": "warmup"}, self.runner("A", LopLengde=150)]}
        self.assertIsNone(validate_setup(valid))
        for broken in (
            dict(valid, type="noe annet"),
            dict(valid, version=2),
            dict(valid, settings=["feil"]),
            dict(valid, items=["ikke en rad"]),
            dict(valid, items=[self.runner("A", LopLengde="3:00")]),
            dict(valid, items=[self.runner("A", LopLengde=-60)]),
            dict(valid, items=[self.runner("A", Påmelding=5)]),
        ):
            self.assertIsNotNone(validate_setup(broken), broken)

    def test_playlist_uses_updated_music_from_zip(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            music_zip = Path(temp_dir) / "Musikk.zip"
            out_dir = Path(temp_dir) / "out"
            rows = [{"MusikkFil": "mappe/sang.mp3", "PrintName": "Test"}]
            for data in (b"ID3-gammel", b"ID3-ny"):
                with zipfile.ZipFile(music_zip, "w") as archive:
                    archive.writestr("mappe/sang.mp3", data)
                self.assertTrue(generate_vlc_playlist(rows, out_dir, "x", music_zip, Mock()))

            self.assertEqual((out_dir / "music" / "mappe" / "sang.mp3").read_bytes(), b"ID3-ny")

    def test_new_playback_stops_running_vlc(self):
        app = App.__new__(App)
        old_process = Mock()
        old_process.poll.return_value = None
        app.vlc_process = old_process
        app.audio_ready = False
        app.external_playback = True

        app.stop_running_playback()

        old_process.terminate.assert_called_once()
        self.assertIsNone(app.vlc_process)
        self.assertFalse(app.external_playback)


if __name__ == "__main__":
    unittest.main()
