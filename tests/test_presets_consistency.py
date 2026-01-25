import os
import sys
from pathlib import Path

import yaml


def test_presets_rule_ids_exist_in_engine():
    project_root = Path(__file__).resolve().parent.parent
    backend_root = project_root / "python-backend"

    sys.path.insert(0, str(backend_root))

    from core.engine import RuleEngine

    presets_path = backend_root / "config" / "presets.yaml"
    assert presets_path.exists(), f"presets.yaml not found at {presets_path}"

    with presets_path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}

    presets = data.get("presets", {}) or {}
    rule_defaults = data.get("rule_defaults", {}) or {}

    referenced_rule_ids = set(rule_defaults.keys())
    for preset in presets.values():
        rules = (preset or {}).get("rules", {}) or {}
        referenced_rule_ids.update(rules.keys())

    engine = RuleEngine()
    engine._load_rules()
    loaded_rule_ids = set(engine.rules.keys())

    missing = sorted(rid for rid in referenced_rule_ids if rid and rid not in loaded_rule_ids)
    assert not missing, f"presets.yaml references unknown rule IDs: {missing}"
