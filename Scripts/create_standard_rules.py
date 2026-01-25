import json
import os
import sys

# Add current directory to path to allow import
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from rule_builder import PRESETS

def create_rules():
    rules = []
    # Add all presets
    print("Generating standard formatting rules...")
    
    # Order matters: Headings first, then Subheadings, then Bullets, then Narrative
    order = ["main_heading", "subheading_a", "subheading_1", "bullet", "narrative"]
    
    for key in order:
        if key in PRESETS:
            print(f"Adding rule: {PRESETS[key]['name']}")
            rules.append(PRESETS[key])
            
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "formatting_rules.json")
    
    with open(output_path, "w") as f:
        json.dump(rules, f, indent=2)
        
    print(f"\nSuccessfully created rules file at: {output_path}")
    print("You can apply these rules using:")
    print(f"python Scripts/rule_engine.py --apply \"<path_to_docx>\" \"{output_path}\"")

if __name__ == "__main__":
    create_rules()
