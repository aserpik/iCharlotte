import json

rules = [
    {
        "name": "Test Replace",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "contains",
            "pattern": "Plaintiffs",
            "case_sensitive": True,
            "whole_word": False
        },
        "action": {
            "type": "replace",
            "replacement": "the Plaintiffs"
        }
    }
]

with open(".gemini/tmp/repro_rules.json", "w") as f:
    json.dump(rules, f)
