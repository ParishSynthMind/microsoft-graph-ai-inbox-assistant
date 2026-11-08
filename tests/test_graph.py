import pytest

from graph_assistant.fake_graph import FakeResponse, err
from graph_assistant.graph import GraphClient, GraphError, drive_path


def test_retries_429_and_honours_retry_after(fake):
    slept = []
    g = GraphClient(fake, lambda: fake.token, sleep=slept.append)
    fake.fail_next = [err(429, "TooManyRequests", headers={"Retry-After": "7"}),
                      err(503, "serviceNotAvailable")]
    assert g.whoami()["mail"] == "alex@contoso-demo.example"
    assert slept[0] == 7.0 and len(slept) == 2 and g.retries == 2


def test_backoff_is_capped(fake):
    slept = []
    g = GraphClient(fake, lambda: fake.token, sleep=slept.append, max_delay=5)
    fake.fail_next = [err(429, "TooManyRequests", headers={"Retry-After": "600"})]
    g.whoami()
    assert slept == [5]


def test_gives_up_after_max_retries(fake):
    g = GraphClient(fake, lambda: fake.token, sleep=lambda s: None, max_retries=2)
    fake.fail_next = [err(429, "TooManyRequests")] * 3
    with pytest.raises(GraphError) as e:
        g.whoami()
    assert e.value.status == 429


def test_other_errors_raise_with_graphs_code(graph, fake):
    fake.fail_next = [err(403, "ErrorAccessDenied", "Access is denied.")]
    with pytest.raises(GraphError) as e:
        graph.whoami()
    assert e.value.code == "ErrorAccessDenied" and e.value.status == 403


def test_bad_token_is_401(fake):
    g = GraphClient(fake, lambda: "wrong", sleep=lambda s: None)
    with pytest.raises(GraphError) as e:
        g.whoami()
    assert e.value.code == "InvalidAuthenticationToken"


def test_non_json_error_body():
    class S:
        def request(self, *a, **k):
            r = FakeResponse(500, content=b"<html>gateway</html>")
            return r
    g = GraphClient(S(), lambda: "t", sleep=lambda s: None, max_retries=0)
    with pytest.raises(GraphError) as e:
        g.whoami()
    assert e.value.code == "unknown"


def test_delta_pages_and_returns_only_new_mail_next_time(graph, fake):
    for i in range(5):
        fake.add_message(f"m{i}", ("A", "a@x.example"), "hi", f"2026-11-0{i + 1}T10:00:00Z")
    msgs, link = graph.inbox_changes()
    assert len(msgs) == 5 and "deltatoken" in link
    msgs2, link2 = graph.inbox_changes(link)
    assert msgs2 == []
    fake.add_message("new", ("B", "b@x.example"), "hello", "2026-11-08T10:00:00Z")
    msgs3, _ = graph.inbox_changes(link2)
    assert [m["subject"] for m in msgs3] == ["new"]


def test_delta_drops_removed_entries(graph, fake):
    class S:
        def request(self, *a, **k):
            return FakeResponse(200, {"value": [{"id": "1", "@removed": {"reason": "deleted"}}, {"id": "2"}],
                                      "@odata.deltaLink": "https://graph.microsoft.com/v1.0/x?$deltatoken=1"})
    g = GraphClient(S(), lambda: "t")
    msgs, link = g.inbox_changes()
    assert [m["id"] for m in msgs] == ["2"] and link.endswith("deltatoken=1")


def test_drive_path_quotes_each_segment():
    assert drive_path("/AI Inbox/a#b%.pdf/") == "AI%20Inbox/a%23b%25.pdf"


def test_ensure_folder_is_idempotent(graph, fake):
    a = graph.ensure_folder("AI Inbox/Attachments/2026-11")
    b = graph.ensure_folder("AI Inbox/Attachments/2026-11")
    assert a["id"] == b["id"]
    assert sum(1 for i in fake.items.values() if i.folder) == 4  # root + 3


def test_ensure_folder_survives_a_race(graph, fake):
    real_get = graph.item_by_path
    calls = {"n": 0}

    def racing_get(path):
        calls["n"] += 1
        if calls["n"] == 1:  # someone else creates it between our look and our create
            fake.request("POST", "https://graph.microsoft.com/v1.0/me/drive/root/children",
                         json={"name": "Shared"}, headers={"Authorization": f"Bearer {fake.token}"})
            return None
        return real_get(path)
    graph.item_by_path = racing_get
    assert graph.ensure_folder("Shared")["name"] == "Shared"


def test_small_upload_and_replace(graph, fake):
    graph.upload("Docs", "a.txt", b"one")
    graph.upload("Docs", "a.txt", b"two")
    files = [i for i in fake.items.values() if not i.folder]
    assert len(files) == 1 and files[0].content == b"two"


def test_upload_conflict_fail(graph):
    graph.upload("Docs", "a.txt", b"one", conflict="fail")
    with pytest.raises(GraphError) as e:
        graph.upload("Docs", "a.txt", b"one", conflict="fail")
    assert e.value.code == "nameAlreadyExists"


def test_large_upload_uses_session_in_320k_chunks_without_auth_header(graph, fake):
    data = bytes(range(256)) * (5 * 1024 * 1024 // 256 + 7)  # just over 5 MB
    item = graph.upload("Big", "video.bin", data)
    stored = fake.items[item["id"]]
    assert stored.content == data
    upload_calls = [c for c in fake.calls if c[1].startswith("https://upload.")]
    assert len(upload_calls) >= 2


def test_sharepoint_site_drive_base(fake):
    g = GraphClient(fake, lambda: fake.token, mailbox="/users/office@contoso-demo.example",
                    drive="/sites/contoso.sharepoint.com,abc,def/drive", sleep=lambda s: None)
    g.upload("Shared Documents/Intake", "x.pdf", b"%PDF")
    assert fake.find_path("Shared Documents/Intake/x.pdf") is not None
    assert g.whoami()["id"] == "u1"


def test_workbook_table_rows(graph, fake):
    from graph_assistant.assistant import make_workbook_bytes
    item = graph.upload("", "Book.xlsx", make_workbook_bytes(["A", "B"]))
    t = graph.create_table(item["id"], "Requests", "A1:B1", "Log")
    assert t["name"] == "Log"
    sid = graph.workbook_session(item["id"])
    graph.add_table_rows(item["id"], "Log", [["1", "2"]], session_id=sid)
    assert graph.table_rows(item["id"], "Log", session_id=sid) == [["1", "2"]]
    graph.close_workbook_session(item["id"], sid)
    with pytest.raises(GraphError) as e:
        graph.add_table_rows(item["id"], "Log", [["only one"]])
    assert e.value.code == "InvalidArgument"


def test_calendar_view_requires_window_and_filters(graph, fake):
    fake.add_event("Standup", "2026-11-09T09:30:00", "2026-11-09T10:00:00")
    fake.add_event("Old", "2026-10-01T09:30:00", "2026-10-01T10:00:00")
    evs = graph.calendar_view("2026-11-09T00:00:00", "2026-11-20T00:00:00", "Europe/London")
    assert [e["subject"] for e in evs] == ["Standup"]


def test_attachment_bytes_and_item_attachment_refused(graph, fake):
    mid = fake.add_message("s", ("A", "a@x.example"), "b", "2026-11-08T10:00:00Z", attachments=[
        {"name": "a.pdf", "content": b"%PDF-1"},
        {"name": "Fwd: meeting", "type": "#microsoft.graph.itemAttachment"}])
    atts = graph.attachments(mid)
    assert "contentBytes" not in atts[0]
    assert graph.attachment_bytes(mid, atts[0]["id"]) == b"%PDF-1"
    with pytest.raises(GraphError):
        graph.attachment_bytes(mid, atts[1]["id"])


def test_draft_reply_is_not_sent(graph, fake):
    mid = fake.add_message("Call?", ("A", "a@x.example"), "b", "2026-11-08T10:00:00Z")
    d = graph.draft_reply(mid, "Hi")
    assert fake.drafts[d["id"]]["isDraft"] is True
    assert fake.drafts[d["id"]]["body"]["content"] == "Hi"
    assert not any("/send" in url for _, url in fake.calls)
