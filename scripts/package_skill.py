"""Package the public skill files, excluding local cases and release archives."""

import argparse
import json
import os
from pathlib import Path
import re
import stat
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile


PUBLIC_FILES = ("SKILL.md", "README.md", "TESTING.md", "VERSION", "requirements.txt", ".gitignore", "LICENSE", "LICENSE.md", "NOTICE", "NOTICE.md")
PUBLIC_DIRS = (".github", "agents", "assets", "examples", "references", "scripts", "tests")
EXCLUDED_DIRS = {"_local", "release", "dist", ".git", "__pycache__", ".venv", ".pytest_cache", ".build", "node_modules"}


def release_version(root, version=None):
    version_file = Path(root) / "VERSION"
    configured = version_file.read_text(encoding="utf-8-sig").strip() if version_file.is_file() else None
    version = version if version is not None else configured if configured is not None else "0.1.0"
    if not re.fullmatch(r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-[0-9A-Za-z]+(?:[.-][0-9A-Za-z]+)*)?", version):
        raise ValueError("Version must use a form such as 0.1.0 or 0.1.0-beta.1")
    if configured is not None and version != configured:
        raise ValueError("Version differs from VERSION; update VERSION before preparing a release")
    return version


def _is_link(path):
    info = path.lstat()
    return stat.S_ISLNK(info.st_mode) or bool(
        getattr(info, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0)
    )


def public_files(root):
    """Select files by public entry points; never scan the whole project root."""
    if not (root / "SKILL.md").is_file():
        raise ValueError("Skill root must contain SKILL.md")
    selected = []
    for name in PUBLIC_FILES:
        path = root / name
        if path.exists():
            if _is_link(path) or not path.is_file():
                raise ValueError(f"Unsupported public file: {path}")
            selected.append(path)
    for name in PUBLIC_DIRS:
        directory = root / name
        if not directory.exists():
            continue
        if _is_link(directory) or not directory.is_dir():
            raise ValueError(f"Unsupported public directory: {directory}")
        for current, directories, filenames in os.walk(directory, followlinks=False):
            parent = Path(current)
            directories[:] = [entry for entry in directories if entry not in EXCLUDED_DIRS]
            for entry in directories:
                if _is_link(parent / entry):
                    raise ValueError(f"Linked directories cannot be packaged: {parent / entry}")
            for entry in filenames:
                path = parent / entry
                if path.suffix.lower() in {".pyc", ".pyo"} or path.name.startswith("~$"):
                    continue
                if _is_link(path):
                    raise ValueError(f"Linked files cannot be packaged: {path}")
                selected.append(path)
    return sorted(selected, key=lambda path: path.relative_to(root).as_posix())


def package_skill(root, output=None, version=None):
    root = Path(root).resolve()
    version = release_version(root, version)
    output = Path(output).resolve() if output else root / "release" / f"v{version}" / f"{root.name}-v{version}.zip"
    if output.suffix.lower() != ".zip":
        raise ValueError("Package output must be a .zip file")
    if output.is_relative_to(root) and output.relative_to(root).parts[0] != "release":
        raise ValueError("Package output inside the project must be under release/")
    files = public_files(root)
    output.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(prefix=".package-", suffix=".zip", dir=output.parent)
    os.close(handle)
    temporary = Path(temporary)
    try:
        expected = [(Path(root.name) / path.relative_to(root)).as_posix() for path in files]
        with ZipFile(temporary, "w", ZIP_DEFLATED) as archive:
            for path, name in zip(files, expected):
                archive.write(path, name)
        with ZipFile(temporary) as archive:
            if archive.testzip() is not None or archive.namelist() != expected:
                raise ValueError("Package verification failed")
        os.replace(temporary, output)
    finally:
        if temporary.exists():
            temporary.unlink()
    return {"package": str(output), "version": version, "files": len(files), "bytes": output.stat().st_size}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, help="Defaults to release/v<version>/<skill-name>-v<version>.zip")
    parser.add_argument("--version", help="Confirms the release version; must match VERSION")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    print(json.dumps(package_skill(root, args.output, args.version), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
