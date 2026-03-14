"""Simple configuration reader for properties file.

Config file is resolved in order: (1) same directory as this module (src/config/),
(2) current working directory. Run from project root (e.g. python outlook_mcp.py)
so that src/config/config.properties is found.
"""
# Author: Hemprasad Badgujar

import logging
import os
from typing import Any, List, Optional

logger = logging.getLogger(__name__)


class ConfigReader:
    """Reads configuration from config.properties file."""
    
    def __init__(self, config_file: str = "config.properties") -> None:
        """
        Initialize ConfigReader and load configuration.
        Args:
            config_file (str): Path to config file.
        """
        self.config_file = config_file
        self.config = {}
        self.load_config()
    
    def _resolve_config_path(self) -> Optional[str]:
        """Resolve config file path: module dir first, then CWD."""
        module_dir = os.path.dirname(os.path.abspath(__file__))
        candidate = os.path.join(module_dir, self.config_file)
        if os.path.exists(candidate):
            return candidate
        cwd_path = os.path.join(os.getcwd(), self.config_file)
        if os.path.exists(cwd_path):
            return cwd_path
        # Also check CWD/config subdir (e.g. when run from project root)
        cwd_config = os.path.join(os.getcwd(), "config", self.config_file)
        if os.path.exists(cwd_config):
            return cwd_config
        return None

    def load_config(self) -> None:
        """
        Load configuration from properties file.
        Uses _resolve_config_path() for location; falls back to defaults if not found.
        """
        config_path = self._resolve_config_path()
        if not config_path:
            logger.warning(
                "Config file %s not found (checked module dir and CWD). Using defaults.",
                self.config_file,
            )
            self._set_defaults()
            return

        try:
            with open(config_path, "r", encoding="utf-8", errors="replace") as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    if "=" in line:
                        key, value = line.split("=", 1)
                        key = key.strip()
                        value = value.strip()
                        self.config[key] = self._convert_value(value)
                    else:
                        logger.warning("Invalid line %d in config file: %s", line_num, line)
            logger.info("Loaded configuration from %s", config_path)
            self._validate_config()
        except OSError as e:
            logger.error("Error reading config file %s: %s", config_path, e)
            self._set_defaults()
        except Exception as e:
            logger.exception("Unexpected error loading config: %s", e)
            self._set_defaults()
    
    def _convert_value(self, value: str) -> Any:
        """
        Convert string value to appropriate type.
        Args:
            value (str): Value to convert.
        Returns:
            Any: Converted value.
        """
        if not isinstance(value, str):
            return value
        # Boolean values
        if value.lower() in ('true', 'false'):
            return value.lower() == 'true'
        
        # Integer values
        try:
            return int(value)
        except ValueError:
            pass
        
        # Float values
        try:
            return float(value)
        except ValueError:
            pass
        
        # List values (comma-separated)
        if ',' in value:
            return [item.strip() for item in value.split(',') if item.strip()]
        
        # String values
        return value

    def _validate_config(self) -> None:
        """Validate loaded config: value ranges and required keys. Log warnings only."""
        int_keys = (
            "max_search_results",
            "max_body_chars",
            "max_search_body_chars",
            "personal_retention_months",
            "shared_retention_months",
            "batch_processing_size",
            "max_retry_attempts",
            "connection_timeout_minutes",
        )
        for key in int_keys:
            val = self.config.get(key)
            if val is None:
                continue
            try:
                n = int(val)
                if key == "max_search_results" and (n < 1 or n > 10000):
                    logger.warning("config %s=%s should be between 1 and 10000", key, n)
                elif "retention" in key and (n < 1 or n > 120):
                    logger.warning("config %s=%s should be between 1 and 120 (months)", key, n)
                elif key == "max_retry_attempts" and (n < 0 or n > 10):
                    logger.warning("config %s=%s should be between 0 and 10", key, n)
            except (ValueError, TypeError):
                logger.warning("config %s has invalid integer value: %s", key, val)

    def _set_defaults(self) -> None:
        """
        Set default configuration values.
        """
        self.config = {
            'shared_mailbox_email': '',
            'shared_mailbox_name': 'Shared Mailbox',
            'personal_retention_months': 6,
            'shared_retention_months': 12,
            'max_search_results': 500,
            'max_body_chars': 0,
            'include_sent_items': True,
            'include_deleted_items': False,
            'connection_timeout_minutes': 10,
            'max_retry_attempts': 3,
            'batch_processing_size': 50,
            'parallel_search_workers': 2,
            'search_cache_ttl_seconds': 3600,
            'search_cache_max_entries': 100,
            'profile_search': False,
            'analyze_importance_levels': True,
            'search_all_folders': False,
            'use_folder_traversal': False,
            'use_extended_mapi_login': True,
            "include_timestamps": True,
            "clean_html_content": True,
            "max_recipients_display": 10,
        }

    def get(self, key: str, default: Any = None) -> Any:
        """
        Get configuration value by key.
        Args:
            key (str): Configuration key.
            default: Default value if key not found.
        Returns:
            Any: Configuration value.
        """
        return self.config.get(key, default)
    
    def get_int(self, key: str, default: int = 0) -> int:
        """
        Get configuration value as integer.
        Args:
            key (str): Configuration key.
            default (int): Default integer value.
        Returns:
            int: Configuration value as integer.
        """
        value = self.config.get(key, default)
        try:
            return int(value)
        except (ValueError, TypeError):
            return default
    
    def get_bool(self, key: str, default: bool = False) -> bool:
        """
        Get configuration value as boolean.
        Args:
            key (str): Configuration key.
            default (bool): Default boolean value.
        Returns:
            bool: Configuration value as boolean.
        """
        value = self.config.get(key, default)
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.lower() in ('true', '1', 'yes', 'on')
        return default
    
    def get_list(self, key: str, default: Optional[List[Any]] = None) -> List[Any]:
        """
        Get configuration value as list.
        Args:
            key (str): Configuration key.
            default (List): Default list value.
        Returns:
            List: Configuration value as list.
        """
        if default is None:
            default = []
        value = self.config.get(key, default)
        if isinstance(value, list):
            return value
        if isinstance(value, str):
            return [item.strip() for item in value.split(',') if item.strip()]
        return default
    
    def reload(self) -> None:
        """Reload configuration from file. Use after updating config on disk."""
        self.load_config()

    def show_config(self) -> None:
        """
        Display current configuration.
        """
        print("\nCurrent Configuration:")
        print("=" * 40)
        for key, value in sorted(self.config.items()):
            # Don't show empty email addresses
            if key == 'shared_mailbox_email' and not value:
                print(f"{key}: <not configured>")
            else:
                print(f"{key}: {value}")
        print("=" * 40)


# Global config instance
config = ConfigReader()
