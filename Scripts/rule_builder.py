import os
import sys
import json
import argparse
from gemini_utils import call_gemini_api, clean_json_string, log_event

RULE_SCHEMA_PROMPT = """
You are a legal document formatting expert and a Microsoft Word Object Model specialist. Your task is to translate a user's natural language request into a JSON formatting rule.

### Rule Schema:
{{
    "name": "Short descriptive name",
    "enabled": true,
    "trigger": {{
        "scope": "paragraph" (default) or "all_text",
        "match_type": "contains", "starts_with", "wildcard", or "regex",
        "pattern": "The text or pattern to match",
        "is_list": boolean (optional, true if target must be a list item),
        "property_match": {{
            "PropertyPath": Value
        }},
        "whole_word": false,
        "case_sensitive": false
    }},
    "action": {{
        "type": "format",
        "formatting": {{
            "dynamic_properties": {{
                "PropertyPath": Value
            }}
        }}
    }}
}}

### Guidelines:
- **Trigger:**
    - **Text Match:** Use `pattern` and `match_type` for text content (e.g., "starts with 'Note'").
    - **Property Match:** Use `property_match` to filter by formatting attributes.
        - "Bold paragraphs" -> "Range.Font.Bold": true
        - "Heading 1 style" -> "Style": "Heading 1"
        - "Indented text" -> "LeftIndent": 36 (0.5 in)
    - You can combine both! (e.g., "Bold paragraphs starting with Note").
- **Action:**
    - Use `dynamic_properties` to set the desired formatting.
    - "Make it red" -> "Range.Font.ColorIndex": 6
- **General:**
    - Map requests to precise Word Object Model paths relative to the **Paragraph** object.
    - Use integers for Enums, float for points.
    - Output ONLY the raw JSON. No explanation.

User Request: {description}
"""

PATTERN_GEN_PROMPT = """
Analyze the following examples of text provided by a user. Generate the most accurate match pattern and match_type ("contains", "starts_with", "regex") to capture these examples while avoiding false positives.

Examples:
{examples}

Output ONLY a JSON object:
{{
    "match_type": "string",
    "pattern": "string"
}}
"""

def build_rule_from_description(description, model=None):
    prompt = RULE_SCHEMA_PROMPT.format(description=description)
    target_models = [model] if model else None
    response = call_gemini_api(prompt, models=target_models)
    if response:
        try:
            cleaned = clean_json_string(response)
            return json.loads(cleaned)
        except Exception as e:
            log_event(f"Error parsing AI rule JSON: {e}", level="error")
            return {"error": str(e)}
    return {"error": "AI failed to generate response"}

def build_pattern_from_examples(examples, model=None):
    prompt = PATTERN_GEN_PROMPT.format(examples="\n".join(examples))
    target_models = [model] if model else None
    response = call_gemini_api(prompt, models=target_models)
    if response:
        try:
            cleaned = clean_json_string(response)
            return json.loads(cleaned)
        except Exception as e:
            log_event(f"Error parsing AI pattern JSON: {e}", level="error")
            return {"error": str(e)}
    return {"error": "AI failed to generate response"}

PRESETS = {
    "main_heading": {
        "name": "Format Main Headings",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "regex",
            "pattern": "^(FACTUAL BACKGROUND|PROCEDURAL HISTORY|LEGAL ANALYSIS|INTRODUCTION|CONCLUSION|[A-Z][A-Z ]{3,})$",
            "case_sensitive": True
        },
        "action": {
            "type": "format",
            "formatting": {
                "font_name": "Times New Roman",
                "font_size": 12,
                "font_bold": True,
                "alignment": "left",
                "space_after": 0
            }
        }
    },
    "subheading_a": {
        "name": "Format Subheading Level A",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "regex",
            "pattern": "^[A-Z]\\.\\s+",
            "case_sensitive": True
        },
        "action": {
            "type": "format",
            "formatting": {
                "font_name": "Times New Roman",
                "font_size": 12,
                "font_bold": True,
                "first_line_indent": 0,
                "left_indent": 0.5
            }
        }
    },
    "subheading_1": {
        "name": "Format Subheading Level 1",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "regex",
            "pattern": "^\\d+\\.\\s+",
            "case_sensitive": True
        },
        "action": {
            "type": "format",
            "formatting": {
                "font_name": "Times New Roman",
                "font_size": 12,
                "font_bold": False,
                "font_italic": True,
                "first_line_indent": 0,
                "left_indent": 1.0
            }
        }
    },
    "bullet": {
        "name": "Format Bullet Points",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "regex",
            "pattern": ".*",
            "is_list": True
        },
        "action": {
            "type": "format",
            "formatting": {
                "left_indent": 1.0,
                "first_line_indent": -0.5,
                "space_after": 6
            }
        }
    },
    "narrative": {
        "name": "Format Narrative Text",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "regex",
            "pattern": "^(?!FACTUAL|PROCEDURAL|LEGAL|INTRODUCTION|CONCLUSION|[A-Z]\\.|\\d+\\.|[\\u2022\\-o]).{10,}",
            "case_sensitive": True
        },
        "action": {
            "type": "format",
            "formatting": {
                "font_name": "Times New Roman",
                "font_size": 12,
                "alignment": "left",
                "first_line_indent": 0.5,
                "space_after": 0
            }
        }
    }
}

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--describe", type=str, help="Natural language description of the rule")
    parser.add_argument("--examples", type=str, nargs="+", help="Examples of text to match")
    parser.add_argument("--model", type=str, help="Gemini model to use")
    parser.add_argument("--preset", type=str, choices=list(PRESETS.keys()), help="Generate a rule from a preset template")
    args = parser.parse_args()

    if args.preset:
        print(json.dumps(PRESETS[args.preset], indent=2))
    elif args.describe:
        result = build_rule_from_description(args.describe, model=args.model)
        print(json.dumps(result, indent=2))
    elif args.examples:
        result = build_pattern_from_examples(args.examples, model=args.model)
        print(json.dumps(result, indent=2))
