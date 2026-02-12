#!/usr/bin/env python3
"""CLI for managing ms365-access API keys.

Usage:
    # Create the two initial keys
    python -m app.cli create-initial-keys

    # Create a single key
    python -m app.cli create-key --name "my-key" --permissions read:mail read:calendar

    # List all keys
    python -m app.cli list-keys

    # Revoke a key by name
    python -m app.cli revoke-key --name "my-key"
"""

import argparse
import asyncio
import hashlib
import json
import secrets
import sys
from datetime import datetime

from sqlalchemy import select

from app.database import engine, async_session_maker, Base
from app.models import ApiKey


async def ensure_tables():
    """Create tables if they don't exist."""
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)


async def create_key(name: str, permissions: list[str]) -> tuple[str, ApiKey]:
    """Create an API key and return (raw_key, api_key_row)."""
    await ensure_tables()

    raw_key = secrets.token_urlsafe(48)
    key_hash = hashlib.sha256(raw_key.encode()).hexdigest()

    async with async_session_maker() as session:
        # Check for existing key with same name
        result = await session.execute(select(ApiKey).where(ApiKey.name == name))
        existing = result.scalar_one_or_none()
        if existing:
            print(f"  [SKIP] Key '{name}' already exists (id={existing.id}, active={existing.is_active})")
            return None, existing

        api_key = ApiKey(
            key_hash=key_hash,
            name=name,
            permissions=json.dumps(sorted(permissions)),
        )
        session.add(api_key)
        await session.commit()
        await session.refresh(api_key)

    return raw_key, api_key


async def cmd_create_initial_keys():
    """Create the two standard keys: marvin-full and clawd-readonly."""
    print("Creating initial API keys...\n")

    keys_to_create = [
        {
            "name": "marvin-full",
            "permissions": [
                "read:mail",
                "read:calendar",
                "read:files",
                "write:draft",
                "write:mail",
                "write:calendar",
                "write:files",
                "admin",
            ],
        },
        {
            "name": "clawd-readonly",
            "permissions": [
                "read:mail",
                "read:calendar",
                "read:files",
                "write:draft",
            ],
        },
    ]

    created_keys = []
    for spec in keys_to_create:
        raw_key, api_key = await create_key(spec["name"], spec["permissions"])
        if raw_key:
            created_keys.append((spec["name"], raw_key))
            print(f"  [OK] Created '{spec['name']}' (id={api_key.id})")
        # else: skip message already printed

    if created_keys:
        print("\n" + "=" * 72)
        print("SAVE THESE KEYS NOW -- they cannot be retrieved later!")
        print("=" * 72)
        for name, key in created_keys:
            print(f"\n  {name}:")
            print(f"    {key}")
        print("\n" + "=" * 72)
        print("\nAdd to .env or consumer config as:")
        print('  Authorization: Bearer <key>')
    else:
        print("\nNo new keys created (all already exist).")


async def cmd_create_key(name: str, permissions: list[str]):
    """Create a single API key."""
    raw_key, api_key = await create_key(name, permissions)
    if raw_key:
        print(f"Created key '{name}' (id={api_key.id})")
        print(f"\n  Raw key (save now, shown only once):")
        print(f"    {raw_key}")
        print(f"\n  Permissions: {', '.join(sorted(permissions))}")
    # else: already-exists message printed by create_key


async def cmd_list_keys():
    """List all API keys."""
    await ensure_tables()

    async with async_session_maker() as session:
        result = await session.execute(select(ApiKey).order_by(ApiKey.created_at.desc()))
        keys = result.scalars().all()

    if not keys:
        print("No API keys found.")
        return

    print(f"{'ID':<5} {'Name':<25} {'Active':<8} {'Last Used':<22} Permissions")
    print("-" * 100)
    for k in keys:
        perms = json.loads(k.permissions) if isinstance(k.permissions, str) else k.permissions
        last_used = k.last_used_at.strftime("%Y-%m-%d %H:%M:%S") if k.last_used_at else "never"
        active = "yes" if k.is_active else "NO"
        print(f"{k.id:<5} {k.name:<25} {active:<8} {last_used:<22} {', '.join(perms)}")


async def cmd_revoke_key(name: str):
    """Revoke (deactivate) an API key by name."""
    await ensure_tables()

    async with async_session_maker() as session:
        result = await session.execute(select(ApiKey).where(ApiKey.name == name))
        api_key = result.scalar_one_or_none()

        if not api_key:
            print(f"No API key found with name '{name}'.")
            sys.exit(1)

        if not api_key.is_active:
            print(f"Key '{name}' is already revoked.")
            return

        api_key.is_active = False
        session.add(api_key)
        await session.commit()
        print(f"Key '{name}' has been revoked.")


def main():
    parser = argparse.ArgumentParser(description="ms365-access API key management")
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # create-initial-keys
    subparsers.add_parser("create-initial-keys", help="Create the two standard API keys")

    # create-key
    create_parser = subparsers.add_parser("create-key", help="Create a single API key")
    create_parser.add_argument("--name", required=True, help="Key name")
    create_parser.add_argument(
        "--permissions",
        nargs="+",
        required=True,
        help="Space-separated permissions (e.g. read:mail read:calendar)",
    )

    # list-keys
    subparsers.add_parser("list-keys", help="List all API keys")

    # revoke-key
    revoke_parser = subparsers.add_parser("revoke-key", help="Revoke an API key by name")
    revoke_parser.add_argument("--name", required=True, help="Key name to revoke")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    if args.command == "create-initial-keys":
        asyncio.run(cmd_create_initial_keys())
    elif args.command == "create-key":
        asyncio.run(cmd_create_key(args.name, args.permissions))
    elif args.command == "list-keys":
        asyncio.run(cmd_list_keys())
    elif args.command == "revoke-key":
        asyncio.run(cmd_revoke_key(args.name))


if __name__ == "__main__":
    main()
