#!/usr/bin/env python3
"""
Local Outlook mail agent for macOS.

Reads the Microsoft Outlook for Mac local profile database in read-only mode.
The default summarizer is local and extractive; an external LLM command can be
plugged in with OUTLOOK_AGENT_LLM_CMD.
"""

from __future__ import annotations

import argparse
import dataclasses
import datetime as dt
import html
import json
import os
import re
import shlex
import sqlite3
import subprocess
import sys
import textwrap
import urllib.parse
from pathlib import Path
from typing import Iterable, Sequence


DEFAULT_PROFILE_DATA = (
    "~/Library/Group Containers/UBF8T346G9.Office/Outlook/"
    "Outlook 15 Profiles/Main Profile/Data"
)
DEFAULT_DB = f"{DEFAULT_PROFILE_DATA}/Outlook.sqlite"
RECEIVED_AT_SQL = (
    "CASE WHEN typeof(m.Message_TimeReceived) IN ('integer', 'real') "
    "THEN datetime(m.Message_TimeReceived, 'unixepoch', 'localtime') "
    "ELSE m.Message_TimeReceived END"
)
SENT_AT_SQL = (
    "CASE WHEN typeof(m.Message_TimeSent) IN ('integer', 'real') "
    "THEN datetime(m.Message_TimeSent, 'unixepoch', 'localtime') "
    "ELSE m.Message_TimeSent END"
)


@dataclasses.dataclass(frozen=True)
class Message:
    id: int
    folder: str
    account: str | None
    subject: str
    sender: str
    recipients: str
    preview: str
    time_received: str | None
    time_sent: str | None
    unread: bool
    has_attachment: bool
    data_file: str | None

    def compact(self) -> dict:
        return {
            "id": self.id,
            "folder": self.folder,
            "account": self.account,
            "time_received": self.time_received,
            "unread": self.unread,
            "has_attachment": self.has_attachment,
            "sender": self.sender,
            "subject": self.subject,
            "preview": self.preview,
        }


class OutlookStore:
    def __init__(self, db_path: str | Path | None = None) -> None:
        self.db_path = Path(
            os.path.expanduser(
                str(db_path or os.environ.get("OUTLOOK_DB") or DEFAULT_DB)
            )
        )
        self.data_dir = self.db_path.parent

    def connect(self) -> sqlite3.Connection:
        if not self.db_path.exists():
            raise FileNotFoundError(
                f"Outlook database not found: {self.db_path}\n"
                "Set OUTLOOK_DB to the path of Outlook.sqlite if your profile differs."
            )
        uri = "file:" + urllib.parse.quote(str(self.db_path)) + "?mode=ro"
        con = sqlite3.connect(uri, uri=True)
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA query_only = ON")
        con.execute("PRAGMA busy_timeout = 3000")
        return con

    def doctor(self) -> dict:
        with self.connect() as con:
            mail_count = con.execute("SELECT count(*) FROM Mail").fetchone()[0]
            folder_count = con.execute("SELECT count(*) FROM Folders").fetchone()[0]
            account_count = con.execute("SELECT count(*) FROM AccountsMail").fetchone()[0]
            latest = con.execute(
                "SELECT CASE WHEN typeof(max(Message_TimeReceived)) IN ('integer', 'real') "
                "THEN datetime(max(Message_TimeReceived), 'unixepoch', 'localtime') "
                "ELSE max(Message_TimeReceived) END FROM Mail"
            ).fetchone()[0]
        return {
            "db_path": str(self.db_path),
            "data_dir": str(self.data_dir),
            "mail_count": mail_count,
            "folder_count": folder_count,
            "account_count": account_count,
            "latest_received": latest,
            "read_only": True,
        }

    def list_messages(
        self,
        *,
        limit: int = 10,
        query: str | None = None,
        folder: str | None = None,
        unread: bool = False,
        since: str | None = None,
    ) -> list[Message]:
        clauses = ["m.Message_Hidden = 0", "m.Message_MarkedForDelete = 0"]
        params: list[object] = []

        if unread:
            clauses.append("m.Message_ReadFlag = 0")
        if folder:
            clauses.append("lower(f.Folder_Name) LIKE lower(?)")
            params.append(f"%{folder}%")
        if since:
            clauses.append(f"coalesce({RECEIVED_AT_SQL}, {SENT_AT_SQL}) >= ?")
            params.append(since)
        if query:
            tokens = [t for t in re.findall(r"[\w@.+-]+", query) if len(t) > 1]
            token_clauses = []
            for token in tokens:
                token_clauses.append(
                    "("
                    "lower(coalesce(m.Message_NormalizedSubject, '')) LIKE lower(?) OR "
                    "lower(coalesce(m.Message_Preview, '')) LIKE lower(?) OR "
                    "lower(coalesce(m.Message_SenderList, '')) LIKE lower(?) OR "
                    "lower(coalesce(m.Message_SenderAddressList, '')) LIKE lower(?)"
                    ")"
                )
                like = f"%{token}%"
                params.extend([like, like, like, like])
            if token_clauses:
                clauses.append("(" + " OR ".join(token_clauses) + ")")

        params.append(max(1, min(limit, 200)))
        sql = f"""
            SELECT
                m.Record_RecordID AS id,
                coalesce(f.Folder_Name, '') AS folder,
                a.Account_EmailAddress AS account,
                coalesce(m.Message_NormalizedSubject, m.Message_ThreadTopic, '') AS subject,
                coalesce(m.Message_SenderList, m.Message_SenderAddressList, '') AS sender,
                coalesce(m.Message_DisplayTo, m.Message_RecipientList, '') AS recipients,
                coalesce(m.Message_Preview, '') AS preview,
                {RECEIVED_AT_SQL} AS time_received,
                {SENT_AT_SQL} AS time_sent,
                coalesce(m.Message_ReadFlag, 0) AS read_flag,
                coalesce(m.Message_HasAttachment, 0) AS has_attachment,
                m.PathToDataFile AS data_file
            FROM Mail m
            LEFT JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
            LEFT JOIN AccountsMail a ON a.Record_RecordID = m.Record_AccountUID
            WHERE {" AND ".join(clauses)}
            ORDER BY coalesce(m.Message_TimeReceived, m.Message_TimeSent) DESC
            LIMIT ?
        """
        with self.connect() as con:
            return [self._message_from_row(row) for row in con.execute(sql, params)]

    def get_message(self, message_id: int) -> Message:
        sql = f"""
            SELECT
                m.Record_RecordID AS id,
                coalesce(f.Folder_Name, '') AS folder,
                a.Account_EmailAddress AS account,
                coalesce(m.Message_NormalizedSubject, m.Message_ThreadTopic, '') AS subject,
                coalesce(m.Message_SenderList, m.Message_SenderAddressList, '') AS sender,
                coalesce(m.Message_DisplayTo, m.Message_RecipientList, '') AS recipients,
                coalesce(m.Message_Preview, '') AS preview,
                {RECEIVED_AT_SQL} AS time_received,
                {SENT_AT_SQL} AS time_sent,
                coalesce(m.Message_ReadFlag, 0) AS read_flag,
                coalesce(m.Message_HasAttachment, 0) AS has_attachment,
                m.PathToDataFile AS data_file
            FROM Mail m
            LEFT JOIN Folders f ON f.Record_RecordID = m.Record_FolderID
            LEFT JOIN AccountsMail a ON a.Record_RecordID = m.Record_AccountUID
            WHERE m.Record_RecordID = ?
        """
        with self.connect() as con:
            row = con.execute(sql, (message_id,)).fetchone()
        if row is None:
            raise KeyError(f"No Outlook message with id {message_id}")
        return self._message_from_row(row)

    def message_text(self, message: Message, *, max_chars: int = 12000) -> str:
        parts = [
            f"Subject: {message.subject}",
            f"From: {message.sender}",
            f"To: {message.recipients}",
            f"Received: {message.time_received or message.time_sent or ''}",
            "",
        ]
        body = ""
        if message.data_file:
            path = self.data_dir / message.data_file
            if path.exists():
                body = extract_olk_text(path, max_chars=max_chars)
        if not body:
            body = message.preview
        parts.append(body[:max_chars])
        return "\n".join(parts).strip()

    @staticmethod
    def _message_from_row(row: sqlite3.Row) -> Message:
        return Message(
            id=int(row["id"]),
            folder=row["folder"] or "",
            account=row["account"],
            subject=clean_inline(row["subject"]),
            sender=clean_inline(row["sender"]),
            recipients=clean_inline(row["recipients"]),
            preview=clean_inline(row["preview"]),
            time_received=row["time_received"],
            time_sent=row["time_sent"],
            unread=not bool(row["read_flag"]),
            has_attachment=bool(row["has_attachment"]),
            data_file=row["data_file"],
        )


def clean_inline(value: object) -> str:
    text = "" if value is None else str(value)
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def extract_olk_text(path: Path, *, max_chars: int = 12000) -> str:
    data = path.read_bytes()
    candidates: list[str] = []

    for encoding in ("utf-8", "utf-16le"):
        decoded = data.decode(encoding, errors="ignore")
        candidates.extend(extract_html_or_text(decoded))

    for raw in re.findall(rb"[\x09\x0a\x0d\x20-\x7e]{24,}", data):
        candidates.append(raw.decode("utf-8", errors="ignore"))

    cleaned: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        for block in normalize_text(candidate).split("\n\n"):
            block = block.strip()
            if not looks_like_body_text(block):
                continue
            key = block[:160].lower()
            if key in seen:
                continue
            seen.add(key)
            cleaned.append(block)
            if sum(len(x) for x in cleaned) >= max_chars:
                break
        if sum(len(x) for x in cleaned) >= max_chars:
            break

    return "\n\n".join(cleaned)[:max_chars].strip()


def extract_html_or_text(text: str) -> list[str]:
    chunks: list[str] = []
    html_match = re.search(r"(?is)<(?:!doctype\s+html|html|body)\b.*", text)
    if html_match:
        chunks.append(strip_html(html_match.group(0)))
    chunks.append(text)
    return chunks


def strip_html(value: str) -> str:
    value = re.sub(r"(?is)<(script|style).*?</\1>", " ", value)
    value = re.sub(r"(?i)<br\s*/?>", "\n", value)
    value = re.sub(r"(?i)</(p|div|li|tr|h[1-6])>", "\n", value)
    value = re.sub(r"<[^>]+>", " ", value)
    return html.unescape(value)


def normalize_text(value: str) -> str:
    value = value.replace("\x00", " ")
    value = re.sub(r"\r\n?", "\n", value)
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def looks_like_body_text(value: str) -> bool:
    if len(value) < 40:
        return False
    printable = sum(ch.isprintable() or ch.isspace() for ch in value)
    letters = sum(ch.isalpha() for ch in value)
    if printable / max(len(value), 1) < 0.85:
        return False
    if letters < 20:
        return False
    noisy_markers = (
        "com.microsoft",
        "Content-Type:",
        "Microsoft-MacOutlook",
        "urn:schemas",
        "http://schemas.microsoft.com",
    )
    return not any(marker.lower() in value[:500].lower() for marker in noisy_markers)


def local_summary(text: str, *, max_sentences: int = 4) -> str:
    text = normalize_text(strip_html(drop_message_headers(text)))
    sentences = split_sentences(text)
    if not sentences:
        return "No readable text found."

    stop = {
        "about",
        "after",
        "again",
        "also",
        "because",
        "before",
        "could",
        "from",
        "have",
        "into",
        "just",
        "more",
        "that",
        "their",
        "there",
        "this",
        "with",
        "would",
        "your",
        "subject",
        "received",
    }
    words = [
        w.lower()
        for w in re.findall(r"[A-Za-z][A-Za-z0-9'-]{2,}", text)
        if w.lower() not in stop
    ]
    freq: dict[str, int] = {}
    for word in words:
        freq[word] = freq.get(word, 0) + 1

    scored = []
    for idx, sentence in enumerate(sentences[:80]):
        tokens = re.findall(r"[A-Za-z][A-Za-z0-9'-]{2,}", sentence.lower())
        score = sum(freq.get(token, 0) for token in tokens) / max(len(tokens), 1)
        if idx < 6:
            score += 0.3
        if "?" in sentence:
            score += 0.15
        scored.append((score, idx, sentence))

    picked = sorted(scored, reverse=True)[:max_sentences]
    picked.sort(key=lambda item: item[1])
    return "\n".join(f"- {item[2]}" for item in picked)


def split_sentences(text: str) -> list[str]:
    text = re.sub(r"\s+", " ", text).strip()
    raw = re.split(r"(?<=[.!?])\s+(?=[A-Z0-9])", text)
    return [s.strip() for s in raw if 30 <= len(s.strip()) <= 500]


def drop_message_headers(text: str) -> str:
    lines = text.splitlines()
    while lines and re.match(r"^(Subject|From|To|Received):\s*", lines[0], re.I):
        lines.pop(0)
    while lines and not lines[0].strip():
        lines.pop(0)
    return "\n".join(lines)


def run_llm(prompt: str) -> str | None:
    command = os.environ.get("OUTLOOK_AGENT_LLM_CMD")
    if not command:
        return None
    result = subprocess.run(
        shlex.split(command),
        input=prompt,
        text=True,
        capture_output=True,
        check=False,
        timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or f"LLM command failed: {command}")
    return result.stdout.strip()


def build_email_context(
    store: OutlookStore, messages: Sequence[Message], *, max_chars: int
) -> str:
    context_blocks = []
    for message in messages:
        text = store.message_text(message, max_chars=max_chars)
        context_blocks.append(
            f"Message {message.id}\n"
            f"From: {message.sender}\n"
            f"Date: {message.time_received or message.time_sent or ''}\n"
            f"Folder: {message.folder}\n"
            f"Subject: {message.subject}\n"
            f"{text}"
        )
    return "\n\n---\n\n".join(context_blocks)


def answer_question(
    store: OutlookStore,
    question: str,
    *,
    limit: int,
    folder: str | None,
    unread: bool,
    since: str | None,
    max_chars: int,
) -> tuple[str, list[Message], bool]:
    messages = store.list_messages(
        limit=limit,
        query=question,
        folder=folder,
        unread=unread,
        since=since,
    )
    if not messages:
        return "No matching messages found.", [], False

    prompt = (
        "You are a private local Outlook mail assistant. Answer the user's "
        "question using only the email excerpts below. Cite message ids when "
        "relevant. If the excerpts do not contain enough evidence, say so. "
        "Keep the answer concise and practical.\n\n"
        f"Question: {question}\n\n"
        f"{build_email_context(store, messages, max_chars=max_chars)}"
    )
    answer = run_llm(prompt)
    if answer:
        return answer, messages, True

    lines = ["Possible matches:"]
    for message in messages:
        date_text = message.time_received or message.time_sent or ""
        lines.append(f"- {message.id} | {date_text} | {message.sender} | {message.subject}")
        if message.preview:
            lines.append(f"  {textwrap.shorten(message.preview, width=140, placeholder='...')}")
    lines.append("")
    lines.append("Local chat mode found matching mail but cannot synthesize answers yet.")
    lines.append("Set OUTLOOK_AGENT_LLM_CMD to enable conversational answers.")
    return "\n".join(lines), messages, False


def print_messages(messages: Sequence[Message], *, as_json: bool = False) -> None:
    if as_json:
        print(json.dumps([m.compact() for m in messages], indent=2, ensure_ascii=False))
        return
    for m in messages:
        flags = []
        if m.unread:
            flags.append("unread")
        if m.has_attachment:
            flags.append("attachment")
        flag_text = f" [{' | '.join(flags)}]" if flags else ""
        date_text = m.time_received or m.time_sent or ""
        print(f"{m.id}  {date_text}  {m.folder}{flag_text}")
        print(f"    From: {m.sender}")
        print(f"    Subject: {m.subject}")
        if m.preview:
            print(f"    {textwrap.shorten(m.preview, width=160, placeholder='...')}")
        print()


def cmd_doctor(args: argparse.Namespace) -> int:
    store = OutlookStore(args.db)
    print(json.dumps(store.doctor(), indent=2))
    return 0


def cmd_list(args: argparse.Namespace) -> int:
    store = OutlookStore(args.db)
    messages = store.list_messages(
        limit=args.limit,
        query=args.query,
        folder=args.folder,
        unread=args.unread,
        since=args.since,
    )
    print_messages(messages, as_json=args.json)
    return 0


def cmd_show(args: argparse.Namespace) -> int:
    store = OutlookStore(args.db)
    message = store.get_message(args.id)
    text = store.message_text(message, max_chars=args.max_chars if args.full else 2000)
    print(text)
    return 0


def cmd_summarize(args: argparse.Namespace) -> int:
    store = OutlookStore(args.db)
    if args.id:
        messages = [store.get_message(args.id)]
    else:
        messages = store.list_messages(
            limit=args.limit,
            query=args.query,
            folder=args.folder,
            unread=args.unread,
            since=args.since,
        )

    for message in messages:
        text = store.message_text(message, max_chars=args.max_chars)
        llm_prompt = (
            "Summarize this email in concise bullets. Include action items, dates, "
            "people, and any needed reply.\n\n"
            f"{text}"
        )
        summary = run_llm(llm_prompt) or local_summary(text)
        print(f"{message.id}  {message.time_received or message.time_sent or ''}")
        print(f"From: {message.sender}")
        print(f"Subject: {message.subject}")
        print(summary)
        print()
    return 0


def cmd_ask(args: argparse.Namespace) -> int:
    store = OutlookStore(args.db)
    answer, _, _ = answer_question(
        store,
        args.question,
        limit=args.limit,
        folder=args.folder,
        unread=args.unread,
        since=args.since,
        max_chars=args.max_chars,
    )
    print(answer)
    return 0


def cmd_chat(args: argparse.Namespace) -> int:
    store = OutlookStore(args.db)
    status = store.doctor()
    llm_status = "LLM answers enabled" if os.environ.get("OUTLOOK_AGENT_LLM_CMD") else "retrieval-only mode"
    print("Outlook Agent Chat")
    print(f"Mailbox: {status['mail_count']} messages, latest {status['latest_received']} ({llm_status})")
    print("Ask a question, or use /recent, /unread, /show ID, /summary ID, /help, /quit.")
    print()

    while True:
        try:
            question = input("mail> ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return 0

        if not question:
            continue
        if question in {"/quit", "/exit", "quit", "exit"}:
            return 0
        if question == "/help":
            print("Commands:")
            print("  /recent             show recent messages")
            print("  /unread             show unread messages")
            print("  /show ID            show one message")
            print("  /summary ID         summarize one message")
            print("  /quit               leave chat")
            print("Natural language examples:")
            print("  what needs a reply from me?")
            print("  summarize messages from finance this week")
            print("  find emails about the budget")
            print()
            continue
        if question == "/recent":
            print_messages(store.list_messages(limit=args.limit))
            continue
        if question == "/unread":
            print_messages(store.list_messages(limit=args.limit, unread=True))
            continue

        command_match = re.match(r"^/(show|summary)\s+(\d+)$", question)
        if command_match:
            command, message_id = command_match.group(1), int(command_match.group(2))
            message = store.get_message(message_id)
            text = store.message_text(message, max_chars=args.max_chars)
            if command == "show":
                print(text)
            else:
                llm_prompt = (
                    "Summarize this email in concise bullets. Include action items, "
                    "dates, people, and any needed reply.\n\n"
                    f"{text}"
                )
                print(run_llm(llm_prompt) or local_summary(text))
            print()
            continue

        answer, _, _ = answer_question(
            store,
            question,
            limit=args.limit,
            folder=args.folder,
            unread=args.unread,
            since=args.since,
            max_chars=args.max_chars,
        )
        print(answer)
        print()
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Read and summarize local Outlook for Mac mail on demand."
    )
    parser.add_argument("--db", help="Path to Outlook.sqlite")
    sub = parser.add_subparsers(required=True)

    sub.add_parser("doctor", help="Check local Outlook database access").set_defaults(
        func=cmd_doctor
    )

    list_p = sub.add_parser("list", help="List recent or matching messages")
    add_filters(list_p)
    list_p.add_argument("--json", action="store_true")
    list_p.set_defaults(func=cmd_list)

    show_p = sub.add_parser("show", help="Show one message by local Outlook id")
    show_p.add_argument("id", type=int)
    show_p.add_argument("--full", action="store_true", help="extract full readable text")
    show_p.add_argument("--max-chars", type=int, default=12000)
    show_p.set_defaults(func=cmd_show)

    sum_p = sub.add_parser("summarize", help="Summarize one or more messages")
    add_filters(sum_p)
    sum_p.add_argument("--id", type=int, help="summarize a single message id")
    sum_p.add_argument("--max-chars", type=int, default=12000)
    sum_p.set_defaults(func=cmd_summarize)

    ask_p = sub.add_parser("ask", help="Retrieve messages relevant to a question")
    ask_p.add_argument("question")
    add_filters(ask_p)
    ask_p.add_argument("--max-chars", type=int, default=5000)
    ask_p.set_defaults(func=cmd_ask)

    chat_p = sub.add_parser("chat", help="Start an interactive mail chat")
    add_filters(chat_p)
    chat_p.add_argument("--max-chars", type=int, default=5000)
    chat_p.set_defaults(func=cmd_chat)

    return parser


def add_filters(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--limit", type=int, default=10)
    parser.add_argument("--query", help="search subject, sender, and preview text")
    parser.add_argument("--folder", help="folder name contains this text")
    parser.add_argument("--unread", action="store_true")
    parser.add_argument("--since", help="lower bound date, for example 2026-06-01")


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except (FileNotFoundError, KeyError, RuntimeError, sqlite3.Error) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
