from __future__ import annotations

import ast
import unittest
from pathlib import Path


class StaticSafetyTests(unittest.TestCase):
    def test_no_outlook_mutation_calls(self):
        root = Path(__file__).resolve().parents[1] / "outlook_archiver"
        forbidden = {"Send", "Delete", "Move", "Quit", "MarkAsTask"}
        hits = []
        for path in root.glob("*.py"):
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
            for node in ast.walk(tree):
                if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and node.func.attr in forbidden:
                    hits.append(f"{path.name}:{node.lineno}:{node.func.attr}")
                if isinstance(node, ast.Assign):
                    for target in node.targets:
                        if isinstance(target, ast.Attribute) and target.attr in {"UnRead", "Categories"}:
                            hits.append(f"{path.name}:{node.lineno}:{target.attr}=")
        self.assertEqual(hits, [], "发现 Outlook 修改调用：" + ", ".join(hits))

    def test_no_network_client_imports(self):
        root = Path(__file__).resolve().parents[1] / "outlook_archiver"
        source = "\n".join(path.read_text(encoding="utf-8") for path in root.glob("*.py"))
        for name in ("requests", "urllib.request", "httpx", "aiohttp"):
            self.assertNotIn("import " + name, source)


if __name__ == "__main__":
    unittest.main()
