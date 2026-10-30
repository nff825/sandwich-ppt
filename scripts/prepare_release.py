"""Prepare a versioned ZIP, clean GitHub upload folder and release notes."""

import argparse
import hashlib
import json
from pathlib import Path
import shutil

from package_skill import package_skill, public_files, release_version


def prepare_release(root, version=None):
    root = Path(root).resolve()
    version = release_version(root, version)
    directory = root / "release" / f"v{version}"
    upload = directory / "github-upload"
    files = public_files(root)
    expected = {path.relative_to(root).as_posix() for path in files}
    if upload.exists():
        existing = {path.relative_to(upload).as_posix() for path in upload.rglob("*") if path.is_file()}
        if existing - expected:
            raise ValueError("Upload directory contains files outside the public release selection")
    report = package_skill(root, version=version)
    for source in files:
        destination = upload / source.relative_to(root)
        destination.parent.mkdir(parents=True, exist_ok=True)
        if destination.is_symlink():
            raise ValueError(f"Upload destination cannot be a link: {destination}")
        shutil.copy2(source, destination)
    for source in files:
        if source.read_bytes() != (upload / source.relative_to(root)).read_bytes():
            raise ValueError("Upload content verification failed")
    notes = directory / "RELEASE_NOTES.txt"
    if not notes.exists():
        notes.write_text(
            f"三明治 PPT v{version}\n\n"
            "采用 MIT License。\n\n"
            "- 内置蓝色 16:9 单页内容模板，也支持用户提供的 PPTX 模板。\n"
            "- 按“上方说明—中部图证据—下方结论”组织内容页。\n"
            "- 支持学术主题标题、来源备注、选择性分组和留白检查。\n"
            "- 附可编辑的合成数据示例、PNG 预览与 CSV。\n"
            "- 提供样式检查、发布打包及自动回归测试。\n\n"
            "运行环境需要具备 PPTX 编辑能力；视觉验收需要渲染能力。\n"
            "示例数据为人工合成，不代表真实测量或算法性能。\n",
            encoding="utf-8",
        )
    report.update({
        "upload_directory": str(upload), "release_notes": str(notes),
        "sha256": hashlib.sha256(Path(report["package"]).read_bytes()).hexdigest(),
    })
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--version", help="Confirms the release version; must match VERSION")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    print(json.dumps(prepare_release(root, args.version), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
