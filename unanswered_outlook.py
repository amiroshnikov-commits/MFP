#!/usr/bin/env python3
"""CLI utility to find Outlook inbox emails that have no reply from you yet.

Works on Windows with Outlook desktop installed and configured.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Iterable, List, Optional, Set

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover - only happens outside Windows with pywin32
    win32com = None

# Outlook constants
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_MAIL_ITEM = 43
OL_VERB_REPLY = 102
OL_VERB_REPLY_ALL = 103
OL_VERB_FORWARD = 104


@dataclass
class MailSnapshot:
    subject: str
    sender: str
    received_time: datetime
    conversation_id: str
    entry_id: str
    last_verb_executed: Optional[int]


def _safe_datetime(value) -> datetime:
    if isinstance(value, datetime):
        return value
    return datetime.min


def collect_replied_conversations(sent_items) -> Set[str]:
    replied: Set[str] = set()
    for item in sent_items:
        if getattr(item, "Class", None) != OL_MAIL_ITEM:
            continue
        conversation_id = str(getattr(item, "ConversationID", "") or "").strip()
        if conversation_id:
            replied.add(conversation_id)
    return replied


def extract_mail_snapshot(item) -> Optional[MailSnapshot]:
    if getattr(item, "Class", None) != OL_MAIL_ITEM:
        return None

    conversation_id = str(getattr(item, "ConversationID", "") or "").strip()
    if not conversation_id:
        return None

    received = _safe_datetime(getattr(item, "ReceivedTime", None))
    return MailSnapshot(
        subject=str(getattr(item, "Subject", "(Без темы)") or "(Без темы)"),
        sender=str(getattr(item, "SenderName", "Неизвестный отправитель") or "Неизвестный отправитель"),
        received_time=received,
        conversation_id=conversation_id,
        entry_id=str(getattr(item, "EntryID", "") or ""),
        last_verb_executed=getattr(item, "LastVerbExecuted", None),
    )


def filter_unanswered(
    inbox_items: Iterable[MailSnapshot],
    replied_conversations: Set[str],
    user_aliases: Set[str],
) -> List[MailSnapshot]:
    unanswered: List[MailSnapshot] = []

    for mail in inbox_items:
        sender_norm = mail.sender.lower().strip()

        # Ignore your own emails in Inbox.
        if sender_norm in user_aliases:
            continue

        # If Outlook marks the message as replied/forwarded, skip.
        if mail.last_verb_executed in {OL_VERB_REPLY, OL_VERB_REPLY_ALL, OL_VERB_FORWARD}:
            continue

        if mail.conversation_id in replied_conversations:
            continue

        unanswered.append(mail)

    # Newest first for convenience.
    unanswered.sort(key=lambda m: m.received_time, reverse=True)
    return unanswered


def _resolve_user_aliases(namespace) -> Set[str]:
    aliases: Set[str] = set()

    current_user = str(getattr(namespace.CurrentUser, "Name", "") or "").strip().lower()
    if current_user:
        aliases.add(current_user)

    try:
        account_count = namespace.Accounts.Count
    except Exception:
        account_count = 0

    for index in range(1, account_count + 1):
        account = namespace.Accounts.Item(index)
        display_name = str(getattr(account, "DisplayName", "") or "").strip().lower()
        smtp = str(getattr(account, "SmtpAddress", "") or "").strip().lower()
        if display_name:
            aliases.add(display_name)
        if smtp:
            aliases.add(smtp)

    return aliases


def _iter_restricted_items(folder, days: int):
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    since = datetime.now() - timedelta(days=days)
    restriction = "[ReceivedTime] >= '{}'".format(since.strftime("%m/%d/%Y %H:%M %p"))
    return items.Restrict(restriction)


def find_unanswered(days: int) -> List[MailSnapshot]:
    if win32com is None:
        raise RuntimeError(
            "Не найден модуль pywin32. Установите его: pip install pywin32 (в Windows)."
        )

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(OL_FOLDER_INBOX)
    sent = outlook.GetDefaultFolder(OL_FOLDER_SENT_MAIL)

    user_aliases = _resolve_user_aliases(outlook)

    sent_items = _iter_restricted_items(sent, days)
    replied_conversations = collect_replied_conversations(sent_items)

    inbox_restricted = _iter_restricted_items(inbox, days)
    inbox_snapshots: List[MailSnapshot] = []
    for item in inbox_restricted:
        mail = extract_mail_snapshot(item)
        if mail:
            inbox_snapshots.append(mail)

    return filter_unanswered(inbox_snapshots, replied_conversations, user_aliases)


def print_report(items: List[MailSnapshot], limit: int) -> None:
    visible = items[:limit] if limit > 0 else items
    if not visible:
        print("🎉 Неотвеченных писем не найдено.")
        return

    print(f"Найдено неотвеченных писем: {len(items)}")
    print("=" * 90)
    for idx, mail in enumerate(visible, start=1):
        ts = mail.received_time.strftime("%Y-%m-%d %H:%M")
        print(f"{idx:>3}. [{ts}] {mail.sender}")
        print(f"     Тема: {mail.subject}")
        print(f"     ConversationID: {mail.conversation_id}")
        print("-" * 90)

    if len(visible) < len(items):
        print(f"Показано {len(visible)} из {len(items)}. Используйте --limit 0 для полного списка.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Показать письма в Outlook Inbox, на которые вы еще не ответили."
    )
    parser.add_argument(
        "--days",
        type=int,
        default=30,
        help="Сканировать письма за последние N дней (по умолчанию: 30).",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=50,
        help="Сколько писем вывести (0 = без лимита).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.days <= 0:
        print("--days должен быть положительным числом.")
        return 2

    try:
        unanswered = find_unanswered(days=args.days)
    except Exception as exc:
        print(f"Ошибка: {exc}")
        return 1

    print_report(unanswered, limit=args.limit)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
