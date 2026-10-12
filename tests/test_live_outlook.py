"""Tests for `export --live --date` (no Outlook needed: the COM objects are fakes)."""

import argparse
import csv
import datetime
import io
import os
import pathlib
import sqlite3
import sys
import tempfile
import unittest
from contextlib import redirect_stderr

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "src"))

from pstostparser import backend_outlook as bo  # noqa: E402
from pstostparser import cli  # noqa: E402
from pstostparser import exporter  # noqa: E402

DAY = datetime.date(2026, 10, 12)
MAIL = 43
MAIL_FOLDER = 0
CALENDAR_FOLDER = 1  # olAppointmentItem


class _FakeCollection:
    """Outlook collections are 1-based and expose Count and Item(i)."""

    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def Item(self, i):
        return self._items[i - 1]


class FakeMail:
    """Stands in for an Outlook item. Counts the reads that matter so tests can see what was touched."""

    def __init__(self, subject, received, cls=MAIL, entry_id="", broken_received=False):
        self.Subject = subject
        self.Class = cls
        self.MessageClass = "IPM.Note" if cls == MAIL else "REPORT.IPM.Note.NDR"
        self.EntryID = entry_id or f"id:{subject}"
        self.SenderEmailAddress = "sender@example.com"
        self.SenderName = "Sender"
        self.To = ""
        self.CC = ""
        self.BCC = ""
        self.SentOn = received
        self.Size = 0
        self.ConversationTopic = ""
        self.Attachments = []
        self._received = received
        self._broken_received = broken_received
        self.received_reads = 0
        self.body_reads = 0

    @property
    def ReceivedTime(self):
        self.received_reads += 1
        if self._broken_received:
            raise AttributeError("ReceivedTime")
        return self._received

    @property
    def Body(self):
        self.body_reads += 1
        return "body"

    @property
    def HTMLBody(self):
        self.body_reads += 1
        return "<p>body</p>"


class FakeItems:
    """Stands in for Folder.Items. Sort puts the items in Outlook's order; a folder built with
    fail_sort refuses to sort, like some stores do."""

    def __init__(self, mails, fail_sort=False):
        self._mails = list(mails)
        self._fail_sort = fail_sort
        self.sort_calls = []

    @property
    def Count(self):
        return len(self._mails)

    def Item(self, i):
        return self._mails[i - 1]

    def Sort(self, prop, descending=False):
        self.sort_calls.append((prop, descending))
        if self._fail_sort:
            raise RuntimeError("Sort refused")
        self._mails.sort(key=lambda m: m._received, reverse=descending)


class FakeFolder:
    def __init__(self, name, mails=(), subfolders=(), fail_sort=False, default_type=MAIL_FOLDER):
        self.Name = name
        self.Folders = list(subfolders)
        self.DefaultItemType = default_type
        self._items = FakeItems(mails, fail_sort)

    @property
    def Items(self):
        return self._items


class _FakeNamespace:
    def __init__(self, stores):
        self.Folders = _FakeCollection(stores)


def store_with(*subfolders):
    """A store root named 'Me' holding the given folders."""
    return FakeFolder("Me", [], list(subfolders))


def live_store(*stores):
    return bo._LiveStore(None, _FakeNamespace(list(stores)))


def run_day(store, **kwargs):
    return list(bo.iter_messages_on_date(store, DAY, **kwargs))


class FolderSelectedTests(unittest.TestCase):
    def test_matches_folder_name_ignoring_case(self):
        self.assertTrue(bo.folder_selected("Me/Inbox", ["inbox"]))

    def test_matches_path_under_the_store(self):
        self.assertTrue(bo.folder_selected("Me/Inbox/Clients", ["Inbox/Clients"]))

    def test_matches_full_path(self):
        self.assertTrue(bo.folder_selected("Me/Inbox", ["Me/Inbox"]))

    def test_other_folders_do_not_match(self):
        self.assertFalse(bo.folder_selected("Me/Sent Items", ["Inbox"]))

    def test_leaf_name_does_not_match_a_different_parent(self):
        self.assertFalse(bo.folder_selected("Me/Archive/Clients", ["Inbox/Clients"]))


class WalkFoldersTests(unittest.TestCase):
    def test_skip_names_prunes_the_whole_subtree(self):
        search = FakeFolder("Search Folders", [], [FakeFolder("Unread Mail")])
        root = FakeFolder("Me", [], [FakeFolder("Inbox"), search])
        paths = [p for p, _ in bo.walk_folders(bo._OutlookStore(None, root),
                                               skip_names={"Search Folders"})]
        self.assertEqual(paths, ["Me", "Me/Inbox"])

    def test_without_skip_names_everything_is_still_walked(self):
        search = FakeFolder("Search Folders", [], [FakeFolder("Unread Mail")])
        root = FakeFolder("Me", [], [FakeFolder("Inbox"), search])
        paths = [p for p, _ in bo.walk_folders(bo._OutlookStore(None, root))]
        self.assertEqual(paths, ["Me", "Me/Inbox", "Me/Search Folders", "Me/Search Folders/Unread Mail"])


class LocalTimeTests(unittest.TestCase):
    def test_naive_received_times_pass_through(self):
        t = datetime.datetime(2026, 10, 12, 0, 30)
        self.assertEqual(bo._local_naive(t), t)

    def test_zone_label_is_dropped_not_converted(self):
        # pywin32 labels Outlook's local wall-clock time with a zone that can be wrong (a Singapore PC
        # came back labelled 'GMT Standard Time'), so the label must not shift the time.
        labelled = datetime.datetime(2026, 10, 12, 15, 36, tzinfo=datetime.timezone.utc)
        self.assertEqual(bo._local_naive(labelled), datetime.datetime(2026, 10, 12, 15, 36))


class IterMessagesOnDateTests(unittest.TestCase):
    def setUp(self):
        self.morning = FakeMail("morning", datetime.datetime(2026, 10, 12, 9, 0))
        self.midnight = FakeMail("midnight", datetime.datetime(2026, 10, 12, 0, 0))
        self.late_yesterday = FakeMail("late yesterday", datetime.datetime(2026, 10, 11, 23, 59, 59))
        self.next_midnight = FakeMail("next midnight", datetime.datetime(2026, 10, 13, 0, 0))
        self.bounce = FakeMail("bounce", datetime.datetime(2026, 10, 12, 10, 0), cls=46)

    def test_keeps_mail_received_on_the_day_and_nothing_else(self):
        inbox = FakeFolder("Inbox", [self.morning, self.midnight, self.late_yesterday,
                                     self.next_midnight, self.bounce])
        got = run_day(live_store(store_with(inbox)))
        self.assertEqual(sorted(m["subject"] for _, _, m in got), ["midnight", "morning"])

    def test_walks_each_mail_folder_newest_first(self):
        inbox = FakeFolder("Inbox", [self.morning])
        run_day(live_store(store_with(inbox)))
        self.assertEqual(inbox.Items.sort_calls, [("[ReceivedTime]", True)])

    def test_stops_at_the_day_and_never_reads_older_mail(self):
        older = FakeMail("older", datetime.datetime(2026, 10, 9, 10, 0))
        yesterday = FakeMail("yesterday", datetime.datetime(2026, 10, 11, 10, 0))
        inbox = FakeFolder("Inbox", [older, self.morning, yesterday])
        run_day(live_store(store_with(inbox)))
        self.assertGreater(self.morning.received_reads, 0)  # the walk really ran
        self.assertEqual(yesterday.received_reads, 1)  # first item past the day is read once
        self.assertEqual(older.received_reads, 0)
        self.assertEqual(older.body_reads, 0)

    def test_reads_bodies_only_for_kept_mail(self):
        inbox = FakeFolder("Inbox", [self.morning, self.late_yesterday, self.bounce])
        run_day(live_store(store_with(inbox)))
        self.assertGreater(self.morning.body_reads, 0)
        self.assertEqual(self.late_yesterday.body_reads, 0)
        self.assertEqual(self.bounce.body_reads, 0)

    def test_skips_search_folders(self):
        virtual = FakeMail("virtual copy", datetime.datetime(2026, 10, 12, 8, 0))
        search = FakeFolder("Search Folders", [], [FakeFolder("Unread Mail", [virtual])])
        inbox = FakeFolder("Inbox", [self.morning])
        got = run_day(live_store(store_with(inbox, search)))
        self.assertEqual([m["subject"] for _, _, m in got], ["morning"])
        self.assertEqual(virtual.body_reads, 0)

    def test_non_mail_folders_such_as_calendars_are_not_read(self):
        calendar = FakeFolder("Calendar", [], default_type=CALENDAR_FOLDER)
        inbox = FakeFolder("Inbox", [self.morning])
        got = run_day(live_store(store_with(inbox, calendar)))
        self.assertEqual([p for p, _, _ in got], ["Me/Inbox"])
        self.assertEqual(calendar.Items.sort_calls, [])

    def test_folder_filter_leaves_other_folders_unsorted(self):
        inbox = FakeFolder("Inbox", [self.morning])
        sent = FakeFolder("Sent Items", [FakeMail("sent", datetime.datetime(2026, 10, 12, 7, 0))])
        got = run_day(live_store(store_with(inbox, sent)), folders=["Inbox"])
        self.assertEqual([p for p, _, _ in got], ["Me/Inbox"])
        self.assertEqual(sent.Items.sort_calls, [])

    def test_skip_ids_drops_mail_already_exported(self):
        inbox = FakeFolder("Inbox", [self.morning])
        stats = {}
        got = run_day(live_store(store_with(inbox)), skip_ids={self.morning.EntryID}, stats=stats)
        self.assertEqual(got, [])
        self.assertEqual(stats["skipped"], 1)

    def test_a_folder_that_cannot_be_sorted_is_read_in_full(self):
        stuck = FakeFolder("Stuck", [self.morning, self.late_yesterday], fail_sort=True)
        stats = {}
        got = run_day(live_store(store_with(stuck)), stats=stats)
        self.assertEqual([m["subject"] for _, _, m in got], ["morning"])
        self.assertEqual(self.late_yesterday.received_reads, 1)
        self.assertEqual(stats["outside"], 1)

    def test_an_item_that_cannot_be_read_is_counted_and_skipped(self):
        broken = FakeMail("broken", datetime.datetime(2026, 10, 12, 9, 30), broken_received=True)
        inbox = FakeFolder("Inbox", [broken, self.morning])
        stats = {}
        got = run_day(live_store(store_with(inbox)), stats=stats)
        self.assertEqual([m["subject"] for _, _, m in got], ["morning"])
        self.assertEqual(stats["errors"], 1)

    def test_stats_count_folders_non_mail_and_other_days(self):
        inbox = FakeFolder("Inbox", [self.morning, self.next_midnight, self.bounce])
        stats = {}
        run_day(live_store(store_with(inbox)), stats=stats)
        self.assertEqual(stats["folders"], 2)  # store root + Inbox
        self.assertEqual(stats["outside"], 1)
        self.assertEqual(stats["non_mail"], 1)

    def test_record_carries_received_date(self):
        inbox = FakeFolder("Inbox", [self.morning])
        got = run_day(live_store(store_with(inbox)))
        self.assertEqual(got[0][2]["received_date"], "2026-10-12T09:00:00")


class MessageRecordTests(unittest.TestCase):
    def test_received_date_sits_beside_the_sent_first_date(self):
        mail = FakeMail("x", datetime.datetime(2026, 10, 12, 0, 10))
        mail.SentOn = datetime.datetime(2026, 10, 11, 23, 50)
        rec = bo.message_to_dict("Me/Inbox", mail, include_body=False)
        self.assertEqual(rec["date"], "2026-10-11T23:50:00")
        self.assertEqual(rec["received_date"], "2026-10-12T00:10:00")

    def test_received_date_is_the_local_wall_clock_without_a_zone_label(self):
        mail = FakeMail("x", datetime.datetime(2026, 10, 12, 15, 36, tzinfo=datetime.timezone.utc))
        rec = bo.message_to_dict("Me/Inbox", mail, include_body=False)
        self.assertEqual(rec["received_date"], "2026-10-12T15:36:00")


class ExporterReceivedDateTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.row = {"subject": "x", "date": "2026-10-11T23:50:00", "received_date": "2026-10-12T00:10:00"}

    def test_csv_has_a_received_date_column(self):
        path = os.path.join(self.tmp.name, "out.csv")
        exporter.write_csv([self.row], path)
        with open(path, encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
        self.assertIn("received_date", reader.fieldnames)
        self.assertEqual(rows[0]["received_date"], "2026-10-12T00:10:00")

    def test_sqlite_has_a_received_date_column(self):
        path = os.path.join(self.tmp.name, "out.db")
        exporter.write_sqlite([self.row], path)
        con = sqlite3.connect(path)
        try:
            value = con.execute("SELECT received_date FROM emails").fetchone()[0]
        finally:
            con.close()
        self.assertEqual(value, "2026-10-12T00:10:00")


def _args(**overrides):
    base = dict(input="", backend="auto", live=False, date="", folder=None, after="", before="")
    base.update(overrides)
    return argparse.Namespace(**base)


class LiveArgumentTests(unittest.TestCase):
    def test_live_with_a_date_and_no_path_is_accepted(self):
        self.assertIsNone(cli.live_arg_error(_args(live=True, date="2026-10-12")))

    def test_live_needs_a_date(self):
        self.assertIn("--date", cli.live_arg_error(_args(live=True)))

    def test_live_rejects_an_impossible_date(self):
        self.assertIn("YYYY-MM-DD", cli.live_arg_error(_args(live=True, date="2026-02-30")))

    def test_live_rejects_a_file_path(self):
        self.assertIn("leave out", cli.live_arg_error(_args(live=True, date="2026-10-12", input="x.pst")))

    def test_live_rejects_other_backends(self):
        self.assertIn("outlook backend",
                      cli.live_arg_error(_args(live=True, date="2026-10-12", backend="aspose")))

    def test_live_rejects_after_and_before(self):
        self.assertIn("--date",
                      cli.live_arg_error(_args(live=True, date="2026-10-12", after="2026-10-01")))

    def test_file_export_needs_a_path(self):
        self.assertIn("--live", cli.live_arg_error(_args()))

    def test_date_and_folder_only_apply_with_live(self):
        self.assertIn("--live", cli.live_arg_error(_args(input="x.pst", date="2026-10-12")))
        self.assertIn("--live", cli.live_arg_error(_args(input="x.pst", folder=["Inbox"])))

    def test_file_export_without_live_flags_is_accepted(self):
        self.assertIsNone(cli.live_arg_error(_args(input="x.pst")))


class LiveParserTests(unittest.TestCase):
    def test_export_takes_live_without_a_path(self):
        a = cli.build_parser().parse_args(
            ["export", "--live", "--date", "2026-10-12", "--folder", "Inbox", "--folder", "Sent Items"])
        self.assertTrue(a.live)
        self.assertEqual(a.input, "")
        self.assertEqual(a.date, "2026-10-12")
        self.assertEqual(a.folder, ["Inbox", "Sent Items"])

    def test_info_still_needs_a_path(self):
        with redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            cli.build_parser().parse_args(["info"])

    def test_live_output_base_name_is_dated(self):
        self.assertEqual(cli.live_base_name("2026-10-12"), "outlook-live_2026-10-12")


if __name__ == "__main__":
    unittest.main()
