import logging
import pathlib
import pyperclip
import re
import socket
import sys
import time
import toml
import traceback
import typing
from datetime import datetime
from pywinauto import Application

logger = logging.getLogger(__name__)

"""
Paint.NET Selection to Clipboard

This script monitors the selection information in Paint.NET and copies it to the clipboard.
The selection information is extracted from the StatusBar and is expected to be in the format "Selection top left: X, Y, Width, Height".
It will copy the selection information to the clipboard in the format "X, Y, Width, Height".
"""

__version__ = "1.0.0"  # Major.Minor.Patch


def read_toml(file_path: typing.Union[str, pathlib.Path]) -> dict:
    """
    Read configuration settings from the TOML file.
    """
    file_path = pathlib.Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f'File not found: "{file_path}"')
    config = toml.load(file_path)
    return config


def get_selection_info(pattern, main_window) -> typing.Union[str, None]:
    try:
        # Locate the StatusBar
        status_bar = main_window.child_window(auto_id="statusBar", control_type="StatusBar")

        # Get all Text elements inside the StatusBar
        texts = status_bar.descendants(control_type="Text")

        # Find the one that contains "Selection top left"
        for t in texts:
            txt = t.window_text()
            if "Selection top left" in txt:
                match = pattern.search(txt)
                if match:
                    return ", ".join(match.groups())
        return None
    except Exception as e:
        print("Error:", e)
        return None


def main() -> None:
    # Connect to Paint.NET
    try:
        app = Application(backend="uia").connect(title_re=".*Paint.NET.*")
        main_window = app.window(title_re=".*Paint.NET.*")
    except Exception as e:
        logger.error("Could not find Paint.NET window.", exc_info=True)
        return

    # Regex to extract four numbers
    pattern = re.compile(r"(\d+)[^\d]+(\d+)[^\d]+(\d+)[^\d]+(\d+)")

    logger.info("Monitoring Paint.NET selection info...")
    last_value = ""
    while True:
        value = get_selection_info(pattern, main_window)
        if value and value != last_value:
            pyperclip.copy(value)
            logger.info(f"Copied to clipboard: {value}")
            last_value = value
        time.sleep(0.5)


def format_duration_long(duration_seconds: float) -> str:
    """
    Format duration in a human-friendly way, showing only the two largest non-zero units.
    For durations >= 1s, do not show microseconds or nanoseconds.
    For durations >= 1m, do not show milliseconds.
    """
    ns = int(duration_seconds * 1_000_000_000)
    units = [
        ('y', 365 * 24 * 60 * 60 * 1_000_000_000),
        ('mo', 30 * 24 * 60 * 60 * 1_000_000_000),
        ('d', 24 * 60 * 60 * 1_000_000_000),
        ('h', 60 * 60 * 1_000_000_000),
        ('m', 60 * 1_000_000_000),
        ('s', 1_000_000_000),
        ('ms', 1_000_000),
        ('us', 1_000),
        ('ns', 1),
    ]
    parts = []
    for name, factor in units:
        value, ns = divmod(ns, factor)
        if value:
            parts.append(f'{value}{name}')
        if len(parts) == 2:
            break
    if not parts:
        return "0s"
    return "".join(parts)


def enforce_max_folder_size(log_dir: pathlib.Path, max_bytes: int) -> None:
    """
    Enforce a maximum total size for all logs in the folder.
    Deletes oldest logs until below limit.
    """
    if max_bytes is None:
        return

    files = sorted(
        [f for f in log_dir.glob("*.log*") if f.is_file()],
        key=lambda f: f.stat().st_mtime
    )

    total_size = sum(f.stat().st_size for f in files)

    while total_size > max_bytes and files:
        oldest = files.pop(0)
        try:
            size = oldest.stat().st_size
            oldest.unlink()
            logger.debug(f'Deleted "{oldest}"')
            total_size -= size
        except Exception:
            logger.error(f'Failed to delete "{oldest}"', exc_info=True)
            continue


def setup_logging(
        logger: logging.Logger,
        log_file_path: typing.Union[str, pathlib.Path],
        max_folder_size_bytes: typing.Union[int, None] = None,
        console_logging_level: int = logging.DEBUG,
        file_logging_level: int = logging.DEBUG,
        log_message_format: str = "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s",
        date_format: str = "%Y-%m-%d %H:%M:%S"
) -> None:

    log_file_path = pathlib.Path(log_file_path)
    log_dir = log_file_path.parent
    log_dir.mkdir(parents=True, exist_ok=True)

    logger.handlers.clear()
    logger.setLevel(file_logging_level)

    formatter = logging.Formatter(log_message_format, datefmt=date_format)

    # File Handler
    file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
    file_handler.setLevel(file_logging_level)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_logging_level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    if max_folder_size_bytes is not None:
        enforce_max_folder_size(log_dir, max_folder_size_bytes)


def load_config(file_path: typing.Union[str, pathlib.Path]) -> dict:
    file_path = pathlib.Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f'File not found: "{file_path}"')
    config = read_toml(file_path)
    return config


if __name__ == "__main__":
    error = 0
    try:
        script_name = pathlib.Path(__file__).stem
        config_path = pathlib.Path(f'{script_name}_config.toml')
        # config_path = pathlib.Path("config.toml")
        config = load_config(config_path)

        logging_config = config.get("logging", {})
        console_logging_level = getattr(logging, logging_config.get("console_logging_level", "INFO").upper(), logging.DEBUG)
        file_logging_level = getattr(logging, logging_config.get("file_logging_level", "INFO").upper(), logging.DEBUG)
        log_message_format = logging_config.get("log_message_format", "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s")
        logs_folder_name = logging_config.get("logs_folder_name", "logs")
        max_folder_size_bytes = logging_config.get("max_folder_size", None)

        pc_name = socket.gethostname()
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        log_dir = pathlib.Path(logs_folder_name) / script_name
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file_name = f'{timestamp}_{script_name}_{pc_name}.log'
        log_file_path = log_dir / log_file_name

        setup_logging(
            logger,
            log_file_path,
            max_folder_size_bytes=max_folder_size_bytes,
            console_logging_level=console_logging_level,
            file_logging_level=file_logging_level,
            log_message_format=log_message_format
        )
        start_time = time.perf_counter_ns()
        logger.info(f'Script: "{script_name}" | Version: {__version__} | Host: "{pc_name}"')
        main()
        end_time = time.perf_counter_ns()
        duration = end_time - start_time
        duration = format_duration_long(duration / 1e9)
        logger.info(f'Execution completed in {duration}.')
    except KeyboardInterrupt:
        logger.warning("Operation interrupted by user.")
        error = 130
    except Exception as e:
        logger.warning(f'A fatal error has occurred: {repr(e)}\n{traceback.format_exc()}')
        error = 1
    finally:
        for handler in logger.handlers:
            handler.close()
        logger.handlers.clear()
        input("Press Enter to exit...")
        sys.exit(error)
