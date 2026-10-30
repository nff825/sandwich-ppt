"""Check that release archives exclude local material and preserve prior output on failure."""

import importlib.util
import os
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "package_skill.py"
SPEC = importlib.util.spec_from_file_location("sandwich_package_skill", SCRIPT)
PACKAGER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(PACKAGER)


class PackageTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(dir=os.environ.get("SANDWICH_TEST_TMPDIR"))
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name) / "fixture-skill"
        self.root.mkdir()
        # Windows TEMP may use an 8.3 alias; match the packager's resolved path.
        self.root = self.root.resolve()
        self.write("VERSION", b"0.1.0\n")

    def write(self, relative, data=b"example"):
        path = self.root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return path

    def test_archive_includes_public_files_and_excludes_local_material(self):
        self.write("SKILL.md")
        self.write("LICENSE", b"MIT License")
        self.write("assets/template.pptx")
        self.write("scripts/tool.py")
        self.write(".github/workflows/tests.yml")
        self.write("_local/cases/private-paper.pdf", b"private source")
        self.write("dist/previous.zip")
        self.write("release/v0.0.1/previous.zip")
        self.write("unrelated/private.txt")
        self.write("scripts/__pycache__/tool.pyc")
        self.write("assets/~$template.pptx")
        report = PACKAGER.package_skill(self.root)
        self.assertEqual(Path(report["package"]), self.root / "release/v0.1.0/fixture-skill-v0.1.0.zip")
        with ZipFile(report["package"]) as archive:
            self.assertEqual(set(archive.namelist()), {
                "fixture-skill/SKILL.md", "fixture-skill/LICENSE", "fixture-skill/VERSION", "fixture-skill/assets/template.pptx",
                "fixture-skill/scripts/tool.py", "fixture-skill/.github/workflows/tests.yml",
            })
            self.assertIsNone(archive.testzip())

    def test_invalid_source_preserves_previous_package(self):
        previous = self.write("release/v0.1.0/fixture-skill-v0.1.0.zip", b"previous package")
        with self.assertRaises(ValueError):
            PACKAGER.package_skill(self.root)
        self.assertEqual(previous.read_bytes(), b"previous package")

    def test_output_cannot_replace_a_public_asset(self):
        self.write("SKILL.md")
        asset = self.write("assets/source.zip", b"source asset")
        with self.assertRaises(ValueError):
            PACKAGER.package_skill(self.root, asset)
        self.assertEqual(asset.read_bytes(), b"source asset")

    def test_invalid_version_cannot_escape_release_directory(self):
        self.write("SKILL.md")
        for version in ("../escape", "1/2/3", "v0.1.0", "0.1"):
            with self.subTest(version=version), self.assertRaises(ValueError):
                PACKAGER.package_skill(self.root, version=version)
        self.assertFalse((self.root / "release").exists())

    def test_updated_version_uses_its_own_directory(self):
        self.write("SKILL.md")
        self.write("VERSION", b"0.2.0-beta.1\n")
        report = PACKAGER.package_skill(self.root)
        self.assertEqual(Path(report["package"]), self.root / "release/v0.2.0-beta.1/fixture-skill-v0.2.0-beta.1.zip")
        with ZipFile(report["package"]) as archive:
            self.assertEqual(archive.read("fixture-skill/VERSION"), b"0.2.0-beta.1\n")

    def test_mismatched_version_preserves_existing_release(self):
        self.write("SKILL.md")
        previous = self.write("release/v0.1.0/fixture-skill-v0.1.0.zip", b"previous release")
        with self.assertRaises(ValueError):
            PACKAGER.package_skill(self.root, version="0.2.0")
        self.assertEqual(previous.read_bytes(), b"previous release")
        self.assertFalse((self.root / "release/v0.2.0").exists())


if __name__ == "__main__":
    unittest.main()
