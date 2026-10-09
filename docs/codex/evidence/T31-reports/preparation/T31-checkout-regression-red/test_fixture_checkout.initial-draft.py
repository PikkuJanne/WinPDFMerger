"""Fixture byte pins must survive real Git checkouts; development tooling only."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest


REPO = Path(__file__).resolve().parents[3]


def load_corpus(root: Path):
    source = root / "tools/test/corpus.py"
    specification = importlib.util.spec_from_file_location("checkout_corpus", source)
    module = importlib.util.module_from_spec(specification)
    specification.loader.exec_module(module)
    return module


def git(root: Path, autocrlf: str, *arguments: str) -> None:
    result = subprocess.run(
        ["git", "-c", f"core.autocrlf={autocrlf}", "-c", "core.safecrlf=false",
         "-c", "core.attributesFile=", *arguments],
        cwd=root, stdin=subprocess.DEVNULL, capture_output=True, timeout=30,
        shell=False,
    )
    if result.returncode:
        raise AssertionError(result.stderr.decode("utf-8", errors="replace"))


class FixtureCheckoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temporary = tempfile.TemporaryDirectory()
        cls.addClassCleanup(cls.temporary.cleanup)
        cls.root = Path(cls.temporary.name)
        cls.seed = cls.root / "seed"
        cls.seed.mkdir()
        corpus = load_corpus(REPO)
        catalog = json.loads(corpus.CATALOG.read_text(encoding="utf-8"))
        paths = {".gitattributes", "tools/test/corpus.py", "tests/fixtures/corpus.json"}
        for recipe, manifest in corpus.GROUPS.values():
            paths.add(recipe)
            if manifest:
                paths.add(manifest)
        paths.update("tests/fixtures/numbered/" + row["file"]
                     for row in catalog["groups"]["numbered"]["fixtures"])
        for relative in paths:
            destination = cls.seed / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(REPO / relative, destination)
        # Reproduce an ordinary Windows add, including its newline conversion.
        # All configuration is scoped to this command and this owned repository.
        git(cls.seed, "true", "init", "--template=", "--initial-branch=checkout-test")
        git(cls.seed, "true", "add", "--all")
        git(cls.seed, "true", "-c", "user.name=Fixture checkout regression",
            "-c", "user.email=fixture-checkout@example.invalid",
            "-c", "commit.gpgsign=false", "commit", "--quiet", "-m", "Synthetic fixture bytes")

    def check_checkout(self, autocrlf: str):
        checkout = self.root / ("checkout-" + autocrlf)
        git(self.root, autocrlf, "clone", "--quiet", "--no-local", str(self.seed), str(checkout))
        corpus = load_corpus(checkout)
        corpus.load_catalog()
        for name, expected in corpus.load_recipe("numbered").expected_files().items():
            self.assertEqual(expected, (checkout / "tests/fixtures/numbered" / name).read_bytes(), name)

    def test_fixture_byte_pins_survive_autocrlf_true_checkout(self):
        self.check_checkout("true")

    def test_fixture_byte_pins_survive_autocrlf_false_checkout(self):
        self.check_checkout("false")


if __name__ == "__main__":
    unittest.main()
