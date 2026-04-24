"""User configuration management using XDG standard.

Stores per-user settings at ~/.config/claude-graph/config.json.
Supports migration from CLAUDE.local.md for backward compatibility.
"""

import json
import re
from pathlib import Path

CONFIG_DIR = Path.home() / ".config" / "claude-graph"
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULT_CONFIG = {
    "email": None,
    "timezone": None,
    "default_building": None,
    "preferred_floors": [],
    "favorite_rooms": [],
}


def get_config() -> dict:
    """Load user config, falling back to CLAUDE.local.md if no config file exists."""
    config = dict(DEFAULT_CONFIG)

    if CONFIG_FILE.exists():
        try:
            stored = json.loads(CONFIG_FILE.read_text())
            config.update({k: v for k, v in stored.items() if v is not None})
        except (json.JSONDecodeError, OSError):
            pass
        return config

    # Migrate from CLAUDE.local.md if it exists
    migrated = _parse_claude_local_md()
    if any(v for v in migrated.values()):
        config.update({k: v for k, v in migrated.items() if v is not None})
        # Save migrated config
        set_config(config)

    return config


def set_config(updates: dict) -> dict:
    """Merge updates into config and save.

    Args:
        updates: Key-value pairs to merge into config.

    Returns:
        The full updated config.
    """
    config = dict(DEFAULT_CONFIG)

    # Load existing
    if CONFIG_FILE.exists():
        try:
            stored = json.loads(CONFIG_FILE.read_text())
            config.update({k: v for k, v in stored.items() if v is not None})
        except (json.JSONDecodeError, OSError):
            pass

    # Apply updates
    config.update({k: v for k, v in updates.items() if k in DEFAULT_CONFIG})

    # Save
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(config, indent=2) + "\n")
    return config


def is_configured() -> bool:
    """Check if config exists with at least an email."""
    if not CONFIG_FILE.exists():
        return False
    try:
        config = json.loads(CONFIG_FILE.read_text())
        return bool(config.get("email"))
    except (json.JSONDecodeError, OSError):
        return False


def _parse_claude_local_md() -> dict:
    """Parse preferences from CLAUDE.local.md (migration helper)."""
    prefs = {}
    local_md = Path(__file__).parent.parent / "CLAUDE.local.md"
    if not local_md.exists():
        return prefs

    try:
        content = local_md.read_text()

        match = re.search(r"default_building:\s*(\S+)", content)
        if match:
            prefs["default_building"] = match.group(1)

        match = re.search(r"preferred_floors:\s*\[([^\]]+)\]", content)
        if match:
            floors = [int(f.strip()) for f in match.group(1).split(",") if f.strip().isdigit()]
            prefs["preferred_floors"] = floors

        match = re.search(r"^timezone:\s*(\S+)", content, re.MULTILINE)
        if match:
            prefs["timezone"] = match.group(1)

        match = re.search(r"^email:\s*(\S+)", content, re.MULTILINE)
        if match:
            prefs["email"] = match.group(1)
    except OSError:
        pass

    return prefs
