#!/usr/bin/env python3
"""
Inkcut launcher for Windows.
Ensures Python 3.9.13 is installed, dependencies are present, and starts Inkcut.
Designed to be PyInstaller-friendly (--onefile --windowed).
"""

import ctypes
import logging
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path
from typing import Iterable, Optional, Tuple

try:
    import winreg  # type: ignore[attr-defined]
except ImportError:  # pragma: no cover
    winreg = None  # Fallback placeholder; script only targets Windows.

REQUIRED_VERSION: Tuple[int, int, int] = (3, 9, 13)
REQUIRED_VERSION_STR = ".".join(str(part) for part in REQUIRED_VERSION)
REQUIRED_PACKAGES = [
    {"pip": "pyqt5", "import": "PyQt5"},
    {"pip": "inkcut", "import": "inkcut"},
]
LOG_FILE_NAME = "inkcut_launcher.log"
RELAUNCH_FLAG = "INKCUT_LAUNCHER_RELAUNCHED"
LOGGER = logging.getLogger("inkcut_launcher")


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def setup_logging(log_path: Path) -> None:
    LOGGER.setLevel(logging.INFO)
    LOGGER.propagate = False
    log_format = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )

    file_handler = logging.FileHandler(log_path, mode="a", encoding="utf-8")
    file_handler.setFormatter(log_format)
    LOGGER.addHandler(file_handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_format)
    LOGGER.addHandler(console_handler)


def shutdown_logging() -> None:
    handlers = list(LOGGER.handlers)
    for handler in handlers:
        handler.flush()
        handler.close()
        LOGGER.removeHandler(handler)


def show_error(message: str) -> None:
    try:
        ctypes.windll.user32.MessageBoxW(  # type: ignore[attr-defined]
            None,
            message,
            "Inkcut Launcher",
            0x00000010,  # MB_ICONERROR
        )
    except Exception:
        LOGGER.error("Unable to show message box.")


def parse_version_from_output(output: str) -> Optional[Tuple[int, int, int]]:
    match = re.search(r"Python\s+(\d+)\.(\d+)\.(\d+)", output)
    if not match:
        return None
    return tuple(int(part) for part in match.groups())


def get_version_for_executable(python_path: Path) -> Optional[Tuple[int, int, int]]:
    try:
        completed = subprocess.run(
            [str(python_path), "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            check=True,
        )
    except (subprocess.CalledProcessError, OSError) as exc:
        LOGGER.warning("Failed to query %s: %s", python_path, exc)
        return None
    return parse_version_from_output(completed.stdout.strip())


def locate_python_executable(required_version: Tuple[int, int, int]) -> Optional[Path]:
    candidates = []

    if winreg is not None:
        registry_paths = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Python\PythonCore\3.9\InstallPath"),
            (
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\WOW6432Node\Python\PythonCore\3.9\InstallPath",
            ),
        ]
        for hive, sub_key in registry_paths:
            try:
                with winreg.OpenKey(hive, sub_key) as key:
                    install_path, _ = winreg.QueryValueEx(key, None)
                    path = Path(install_path).joinpath("python.exe")
                    candidates.append(path)
            except OSError:
                continue

    candidates.extend(
        [
            Path(r"C:\Program Files\Python39\python.exe"),
            Path(r"C:\Program Files (x86)\Python39\python.exe"),
        ]
    )

    local_app = os.environ.get("LOCALAPPDATA")
    if local_app:
        candidates.append(Path(local_app).joinpath("Programs", "Python", "Python39", "python.exe"))

    for candidate in candidates:
        if candidate and candidate.exists():
            version = get_version_for_executable(candidate)
            if version == required_version:
                return candidate

    try:
        completed = subprocess.run(
            [
                "py",
                f"-{required_version[0]}.{required_version[1]}",
                "-c",
                "import sys; print(sys.executable)",
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            check=True,
        )
        potential_path = Path(completed.stdout.strip())
        if potential_path.exists():
            version = get_version_for_executable(potential_path)
            if version == required_version:
                return potential_path
    except (subprocess.CalledProcessError, FileNotFoundError):
        pass

    return None


def is_user_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())  # type: ignore[attr-defined]
    except Exception:
        return False


def determine_installer_filename(version_str: str) -> str:
    machine = platform.machine().lower()
    suffix = ""
    if "arm64" in machine or "aarch64" in machine:
        suffix = "-arm64"
    elif "64" in machine:
        suffix = "-amd64"
    return f"python-{version_str}{suffix}.exe"


def build_python_download_urls(version_str: str) -> Iterable[str]:
    filename = determine_installer_filename(version_str)
    mirrors = [
        "https://www.python.org/ftp/python/{version}/{filename}",
        "https://download.python.org/ftp/python/{version}/{filename}",
        "https://mirrors.huaweicloud.com/python/{version}/{filename}",
    ]
    seen = set()
    for template in mirrors:
        url = template.format(version=version_str, filename=filename)
        if url not in seen:
            seen.add(url)
            yield url


def download_python_installer(urls: Iterable[str], destination: Path) -> Path:
    failures = []
    for url in urls:
        LOGGER.info("Downloading Python %s from %s", REQUIRED_VERSION_STR, url)
        temp_path = destination.with_suffix(".download")
        last_percent = -1
        try:
            with urllib.request.urlopen(url) as response, open(temp_path, "wb") as file:
                total = int(response.getheader("Content-Length", "0"))
                downloaded = 0
                chunk_size = 8192
                while True:
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                    file.write(chunk)
                    downloaded += len(chunk)
                    if total:
                        percent = int(downloaded * 100 / total)
                        if percent != last_percent and percent % 5 == 0:
                            LOGGER.info(
                                "Downloading Python installer: %d%% (%d/%d KB)",
                                percent,
                                downloaded // 1024,
                                total // 1024,
                            )
                            last_percent = percent
            temp_path.replace(destination)
            LOGGER.info("Python installer saved to %s", destination)
            return destination
        except Exception as exc:
            failures.append((url, exc))
            LOGGER.warning("Download from %s failed: %s", url, exc)
            if temp_path.exists():
                temp_path.unlink(missing_ok=True)
    messages = ", ".join(f"{url} ({error})" for url, error in failures)
    raise RuntimeError(f"Failed to download Python installer from all mirrors: {messages}")


def run_python_installer(installer_path: Path) -> None:
    LOGGER.info("Installing Python %s...", REQUIRED_VERSION_STR)
    installer_args = "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_launcher=0"
    if is_user_admin():
        try:
            subprocess.run(
                [
                    str(installer_path),
                    "/quiet",
                    "InstallAllUsers=1",
                    "PrependPath=1",
                    "Include_test=0",
                    "Include_launcher=0",
                ],
                check=True,
            )
        except subprocess.CalledProcessError as exc:
            raise RuntimeError(f"Python installer exited with code {exc.returncode}") from exc
    else:
        command = (
            f"$p = Start-Process -FilePath \"{installer_path}\" "
            f"-ArgumentList '{installer_args}' -Verb RunAs -Wait -PassThru; "
            "exit $p.ExitCode"
        )
        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    command,
                ],
                check=True,
            )
        except subprocess.CalledProcessError as exc:
            if exc.returncode == 1223:
                raise RuntimeError("Python installation canceled by user.") from exc
            raise RuntimeError(f"Elevated installer failed with code {exc.returncode}.") from exc
    LOGGER.info("Python installation finished.")


def prompt_for_local_installer(expected_filename: str) -> Optional[Path]:
    message = (
        "Unable to download Python 3.9.13 automatically.\n"
        f"Click OK and select {expected_filename}."
    )
    try:
        ctypes.windll.user32.MessageBoxW(  # type: ignore[attr-defined]
            None,
            message,
            "Inkcut Launcher",
            0x00000040,  # MB_ICONINFORMATION
        )
    except Exception:
        LOGGER.info("Message box unavailable; continuing with file selection dialog.")

    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:
        LOGGER.error("Tkinter is unavailable for manual installer selection: %s", exc)
        return None

    root = tk.Tk()
    root.withdraw()
    try:
        selected = filedialog.askopenfilename(
            title="Select Python installer",
            filetypes=[("Python Installer", expected_filename), ("Executables", "*.exe")],
        )
    finally:
        root.destroy()

    if not selected:
        LOGGER.warning("User canceled installer selection.")
        return None

    installer_path = Path(selected)
    if not installer_path.exists():
        LOGGER.error("Installer %s does not exist.", installer_path)
        return None

    if installer_path.name.lower() != expected_filename.lower():
        LOGGER.error("Selected %s, expected %s.", installer_path.name, expected_filename)
        return None

    return installer_path


def wait_for_python_install(required_version: Tuple[int, int, int], timeout_seconds: int = 300) -> Optional[Path]:
    LOGGER.info("Validating Python installation...")
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        found = locate_python_executable(required_version)
        if found:
            LOGGER.info("Validated Python at %s", found)
            return found
        time.sleep(5)
    return None


def install_python() -> Path:
    base_temp_dir = Path(tempfile.mkdtemp(prefix="inkcut_python_"))
    installer_filename = determine_installer_filename(REQUIRED_VERSION_STR)
    downloaded_installer_path = base_temp_dir.joinpath(installer_filename)
    installer_to_use: Optional[Path] = None
    try:
        download_urls = list(build_python_download_urls(REQUIRED_VERSION_STR))
        try:
            installer_to_use = download_python_installer(download_urls, downloaded_installer_path)
        except RuntimeError as download_error:
            LOGGER.error("Automatic download failed for all mirrors: %s", download_error)
            installer_to_use = prompt_for_local_installer(installer_filename)
            if not installer_to_use:
                raise RuntimeError("Python installer was not provided by the user.") from download_error
        run_python_installer(installer_to_use)
        installed_path = wait_for_python_install(REQUIRED_VERSION)
        if not installed_path:
            raise RuntimeError("Python installation was not detected after setup.")
        return installed_path
    finally:
        try:
            if downloaded_installer_path.exists():
                downloaded_installer_path.unlink()
        except Exception:
            LOGGER.warning("Unable to remove installer %s", downloaded_installer_path)
        shutil.rmtree(base_temp_dir, ignore_errors=True)


def relaunch_with_python(python_executable: Path) -> None:
    LOGGER.info("Restarting launcher with %s", python_executable)
    shutdown_logging()
    env = os.environ.copy()
    env[RELAUNCH_FLAG] = "1"
    command = [str(python_executable)] + sys.argv
    try:
        completed = subprocess.run(command, env=env)
    except Exception as exc:
        show_error(f"Failed to relaunch with Python 3.9.13: {exc}")
        sys.exit(1)
    sys.exit(completed.returncode)


def ensure_python_environment() -> Path:
    running_frozen = getattr(sys, "frozen", False)
    current_version = sys.version_info[:3]
    current_python = Path(sys.executable).resolve()
    existing = locate_python_executable(REQUIRED_VERSION)

    if not running_frozen:
        if current_version == REQUIRED_VERSION:
            LOGGER.info("Required Python version already active: %s", REQUIRED_VERSION_STR)
            return current_python

        LOGGER.warning(
            "Python %s is required, but current interpreter is %s.%s.%s.",
            REQUIRED_VERSION_STR,
            current_version[0],
            current_version[1],
            current_version[2],
        )

        if existing:
            if current_python != existing.resolve():
                LOGGER.info("Found required Python at %s; relaunching.", existing)
                relaunch_with_python(existing)
            return existing

        installed_path = install_python()
        relaunch_with_python(installed_path)
        return installed_path

    if current_version != REQUIRED_VERSION:
        LOGGER.warning(
            "Bundled Python %s.%s.%s detected; Python %s is required for package management.",
            current_version[0],
            current_version[1],
            current_version[2],
            REQUIRED_VERSION_STR,
        )

    if existing:
        LOGGER.info("Using Python interpreter at %s", existing)
        return existing

    LOGGER.info("Python %s is not installed; proceeding with installation.", REQUIRED_VERSION_STR)
    installed_path = install_python()
    LOGGER.info("Using Python interpreter at %s", installed_path)
    return installed_path


def is_package_installed(import_name: str, python_path: Path) -> bool:
    try:
        subprocess.run(
            [str(python_path), "-c", f"import {import_name}"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True,
        )
        return True
    except subprocess.CalledProcessError:
        return False
    except FileNotFoundError as exc:
        raise RuntimeError(f"Python interpreter {python_path} not found.") from exc


def stream_command_output(command: Iterable[str]) -> int:
    process = subprocess.Popen(
        list(command),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        bufsize=1,
    )
    assert process.stdout is not None
    for raw_line in process.stdout:
        # Split on carriage returns to surface pip progress updates.
        for segment in raw_line.rstrip().split("\r"):
            if segment:
                LOGGER.info(segment)
    process.stdout.close()
    return process.wait()


def install_package(python_path: Path, package_name: str) -> None:
    LOGGER.info("Installing package %s via pip...", package_name)
    command = [
        str(python_path),
        "-m",
        "pip",
        "install",
        package_name,
        "--disable-pip-version-check",
    ]
    return_code = stream_command_output(command)
    if return_code != 0:
        raise RuntimeError(f"pip exited with code {return_code} while installing {package_name}")
    LOGGER.info("Package %s installed successfully.", package_name)


def ensure_packages(python_path: Path, packages: Iterable[dict]) -> None:
    package_list = list(packages)
    total = len(package_list)
    for index, pkg in enumerate(package_list, start=1):
        pip_name = pkg["pip"]
        import_name = pkg["import"]
        LOGGER.info("[%d/%d] Checking %s", index, total, pip_name)
        if is_package_installed(import_name, python_path):
            LOGGER.info("Package %s already present.", pip_name)
            continue
        install_package(python_path, pip_name)


def launch_inkcut(python_path: Path) -> None:
    LOGGER.info("Launching Inkcut...")
    try:
        process = subprocess.Popen(
            [str(python_path), "-m", "inkcut"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:
        raise RuntimeError(f"Failed to start Inkcut: {exc}") from exc

    time.sleep(2)
    exit_code = process.poll()
    if exit_code is None:
        LOGGER.info("Inkcut started successfully (PID %s).", process.pid)
    else:
        raise RuntimeError(f"Inkcut exited immediately with code {exit_code}")


def main() -> None:
    base_dir = get_base_dir()
    log_path = base_dir.joinpath(LOG_FILE_NAME)
    setup_logging(log_path)
    LOGGER.info("Inkcut launcher started (Python %s).", sys.version.replace("\n", " "))

    try:
        python_path = ensure_python_environment()
        LOGGER.info("Python interpreter selected: %s", python_path)
        ensure_packages(python_path, REQUIRED_PACKAGES)
        launch_inkcut(python_path)
        LOGGER.info("Inkcut launcher finished successfully.")
    except Exception as exc:
        LOGGER.exception("Launcher failed: %s", exc)
        show_error(f"Inkcut launcher error:\n{exc}")
        sys.exit(1)
    finally:
        shutdown_logging()


if __name__ == "__main__":
    main()
