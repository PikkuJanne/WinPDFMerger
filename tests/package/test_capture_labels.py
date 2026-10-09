"""Regression for Windows capture paths; never import execution orchestration."""
import ast
from pathlib import Path
import unittest

source = Path(__file__).resolve().parents[2] / 'docs/codex/evidence/T28-reports/scripts/capture-T28.py'
tree = ast.parse(source.read_text(encoding='utf-8'))
helper = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == 'validate_capture_labels')
namespace = {}
exec(compile(ast.Module(body=[helper], type_ignores=[]), str(source), 'exec'), namespace)
validate = namespace['validate_capture_labels']


class CapturePaths(unittest.TestCase):
    def test_refuses_original_windows_static_collision(self):
        with self.assertRaises(ValueError):
            validate(['PS51-Static', 'PS51-static'])

    def test_refuses_duplicate_names(self):
        with self.assertRaises(ValueError):
            validate(['PS7-build-first', 'PS7-build-first'])

    def test_accepts_independent_tier_analyzer_and_build_paths(self):
        validate(['PS51-Static', 'PS51-Static-export', 'PS51-selected-static',
                  'PS51-build-first', 'PS51-build-repeat', 'PS7-Static',
                  'PS7-Static-export', 'PS7-selected-static'])


if __name__ == '__main__':
    unittest.main()
