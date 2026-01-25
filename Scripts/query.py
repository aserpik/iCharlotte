import sys
import os
import json
import argparse

# Ensure Scripts/ is in path
sys.path.append(os.path.join(os.getcwd(), 'Scripts'))
from case_data_manager import CaseDataManager

def main():
    parser = argparse.ArgumentParser(description="Query case data by tag.")
    parser.add_argument("file_number", help="The case file number (e.g., 5800.015)")
    parser.add_argument("tag", help="The tag to search for (e.g., 'Medical', 'Complaint')")
    parser.add_argument("--keys-only", action="store_true", help="Show only variable names, not values")
    
    args = parser.parse_args()
    
    manager = CaseDataManager()
    
    print(f"--- Querying {args.file_number} for tag: '{args.tag}' ---")
    
    # get_by_tag returns {key: value} by default in the current implementation, 
    # but let's look at the raw objects to check tags explicitly if needed.
    # Actually, CaseDataManager.get_by_tag returns values. 
    # Let's inspect get_all_variables to allow partial tag matching.
    
    all_vars = manager.get_all_variables(args.file_number, flatten=False)
    
    matches = {}
    for key, data in all_vars.items():
        if not isinstance(data, dict) or "tags" not in data:
            continue
            
        # Case-insensitive partial match
        tags = [t.lower() for t in data["tags"]]
        search_tag = args.tag.lower()
        
        if any(search_tag in t for t in tags):
            matches[key] = data

    if not matches:
        print("No matches found.")
        return

    print(f"Found {len(matches)} matches:\n")
    
    for key, data in matches.items():
        print(f"Variable: {key}")
        print(f"Tags: {data['tags']}")
        if not args.keys_only:
            val = str(data['value'])
            # Truncate long values for display
            if len(val) > 200:
                val = val[:200] + "... [truncated]"
            print(f"Value: {val}")
        print("-" * 40)

if __name__ == "__main__":
    main()
