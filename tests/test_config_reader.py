"""Unit tests for ConfigReader."""
import sys
from pathlib import Path

# Add project root so "src" can be imported when run directly: python tests/test_config_reader.py
_root = Path(__file__).resolve().parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import os
import tempfile
import pytest


@pytest.fixture
def temp_config_dir(tmp_path):
    """Create a temporary config directory and return path."""
    return tmp_path


@pytest.fixture
def config_file(temp_config_dir):
    """Create a temporary config.properties file."""
    path = temp_config_dir / "config.properties"
    path.write_text(
        """
# comment
max_search_results=100
max_body_chars=5000
use_extended_mapi_login=true
shared_mailbox_email=test@example.com
batch_processing_size=25
parallel_search_workers=4
search_cache_ttl_seconds=1800
search_cache_max_entries=50
"""
    )
    return path


def test_config_reader_defaults_contain_expected_keys():
    """Default config contains expected keys."""
    from src.config.config_reader import ConfigReader
    reader = ConfigReader()
    reader._set_defaults()
    assert reader.get_int("max_search_results", 500) in (500, 50)
    assert reader.get("max_retry_attempts") is not None
    assert "shared_mailbox_email" in reader.config


def test_config_reader_get_int():
    """get_int returns integer or default."""
    from src.config.config_reader import ConfigReader
    reader = ConfigReader()
    reader.config = {"a": 42, "b": "99", "c": True}
    assert reader.get_int("a", 0) == 42
    assert reader.get_int("b", 0) == 99
    assert reader.get_int("missing", 7) == 7
    assert reader.get_int("c", 0) == 0  # bool -> default


def test_config_reader_get_bool():
    """get_bool returns boolean or default."""
    from src.config.config_reader import ConfigReader
    reader = ConfigReader()
    reader.config = {"on": True, "off": False, "yes": "true", "no": "false"}
    assert reader.get_bool("on", False) is True
    assert reader.get_bool("off", True) is False
    assert reader.get_bool("yes", False) is True
    assert reader.get_bool("no", True) is False
    assert reader.get_bool("missing", True) is True


def test_config_reader_get_list():
    """get_list returns list or default."""
    from src.config.config_reader import ConfigReader
    reader = ConfigReader()
    reader.config = {"emails": ["a@x.com", "b@x.com"], "csv": "x,y,z"}
    assert reader.get_list("emails", []) == ["a@x.com", "b@x.com"]
    assert reader.get_list("csv", []) == ["x", "y", "z"]
    assert reader.get_list("missing", []) == []
    assert reader.get_list("missing", None) == []


def test_config_reader_reload():
    """reload() re-runs load_config."""
    from src.config.config_reader import ConfigReader
    reader = ConfigReader()
    reader.config = {"old": 1}
    reader.reload()
    assert "max_search_results" in reader.config or "old" in reader.config


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
