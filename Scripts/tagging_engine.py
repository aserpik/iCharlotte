import os
import json
import logging
import re
from typing import List, Dict, Any, Optional

# Import shared utilities
try:
    from gemini_utils import call_gemini_api, clean_json_string, log_event
except ImportError:
    from Scripts.gemini_utils import call_gemini_api, clean_json_string, log_event

class TaggingEngine:
    def __init__(self, rules_path: str = None):
        if rules_path:
            self.rules_path = rules_path
        else:
            self.rules_path = os.path.join(os.getcwd(), ".gemini", "config", "tagging_rules.json")
            
        self.rules_tree = self._load_rules()

    def _load_rules(self) -> List[Dict]:
        if not os.path.exists(self.rules_path):
            log_event(f"Tagging rules not found at {self.rules_path}.", level="warning")
            return []
        try:
            with open(self.rules_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            log_event(f"Error loading tagging rules: {e}", level="error")
            return []

    def _evaluate_node(self, node: Dict, value_str: str, context_desc: str, known_parties: List[str] = None) -> List[str]:
        """
        Recursively processes a node and its subtags.
        """
        tag_name = node.get("tag")
        description = node.get("description", "")
        action = node.get("action")
        multi_select = node.get("multi_select", False)
        flatten = node.get("flatten", False)
        subtags = node.get("subtags", [])

        collected_tags = []

        # 1. Handle "extract_value" Action
        if action == "extract_value":
            parties_context = ""
            instruction_suffix = ""
            if known_parties:
                # Ensure known_parties contains unique, non-empty strings
                clean_parties = list(set([p for p in known_parties if p and isinstance(p, str)]))
                if clean_parties:
                    parties_context = "\n[KNOWN PARTIES]\n" + ", ".join(clean_parties)
                    instruction_suffix = "\nIf the extracted value matches one of the Known Parties (even with slight variation), RETURN THE EXACT SPELLING FROM THE KNOWN PARTIES LIST."

            prompt = f"""
            Extract a specific value from the text below based on the following instruction:
            Instruction: {description}
            
            [TEXT]
            Context: {context_desc}
            Content: "{value_str}"{parties_context}
            
            Return ONLY the extracted string (e.g. "City of Hesperia"). {instruction_suffix}
            If no specific value can be found, return "Unknown".
            """
            extracted = call_gemini_api(prompt, models=["gemini-3-flash-preview"])
            if extracted and "UNKNOWN" not in extracted.upper():
                val = extracted.strip().replace('"', '').replace("'", "")
                # Apply tag name as prefix if requested (e.g. "Propounding Party: City of Hesperia")
                if tag_name:
                    val = f"{tag_name}: {val}"
                return [val]
            return []

        # 2. Add current tag to collection (unless flattened)
        if tag_name and not flatten:
            collected_tags.append(tag_name)

        # 3. Process Subtags
        if subtags:
            if multi_select:
                # Iterate through all branches and collect ALL that apply
                for sub in subtags:
                    res = self._evaluate_node(sub, value_str, context_desc, known_parties=known_parties)
                    collected_tags.extend(res)
            else:
                # Pick the SINGLE BEST subtag from the list
                options = [{"tag": s["tag"], "desc": s["description"]} for s in subtags]
                prompt = f"""
                Select the SINGLE BEST sub-category for the content provided.
                
                [OPTIONS]
                {json.dumps(options, indent=2)}
                
                [CONTENT]
                Context: {context_desc}
                Value: "{value_str}"
                
                Return ONLY the tag name or "NONE".
                """
                selection = call_gemini_api(prompt, models=["gemini-3-flash-preview"])
                if selection:
                    selected_tag_str = selection.strip().replace('"', '').replace("'", "")
                    selected_node = next((s for s in subtags if s["tag"] == selected_tag_str), None)
                    if selected_node:
                        res = self._evaluate_node(selected_node, value_str, context_desc, known_parties=known_parties)
                        collected_tags.extend(res)

        return collected_tags

    def generate_tags(self, value: Any, context_description: str = "", known_parties: List[str] = None) -> List[str]:
        if not value or str(value).lower() == "none":
            return []
            
        value_str = str(value)
        if len(value_str) > 10000:
            value_str = value_str[:10000] + "...[TRUNCATED]"

        if not self.rules_tree:
            return []

        # Find the Top-Level Match
        top_options = [{"tag": r["tag"], "desc": r["description"]} for r in self.rules_tree]
        prompt = f"""
        Classify this legal variable into ONE of these top-level categories.
        
        [CATEGORIES]
        {json.dumps(top_options, indent=2)}
        
        [VARIABLE]
        Description: {context_description}
        Value: "{value_str}"
        
        Return ONLY the tag name (e.g. "Pleading") or "NONE".
        """
        
        response = call_gemini_api(prompt, models=["gemini-3-flash-preview"])
        if response:
            top_tag = response.strip().replace('"', '').replace("'", "")
            root_node = next((r for r in self.rules_tree if r["tag"] == top_tag), None)
            
            if root_node:
                # Start recursive descent
                return self._evaluate_node(root_node, value_str, context_description, known_parties=known_parties)

        return []

if __name__ == "__main__":
    # Test stub
    engine = TaggingEngine()
    test_val = "Defendant City of Hesperia's Responses to Plaintiff's Form Interrogatories, Set One."
    # Simulated known parties
    known_parties = ["City of Hesperia", "Plaintiff Jane Doe"]
    print(f"Testing with: {test_val}")
    print(f"Known Parties: {known_parties}")
    print(f"Result: {engine.generate_tags(test_val, 'Discovery Document', known_parties)}")
