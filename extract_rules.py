import re

def extract_rules(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Split by file sections
    sections = content.split('==============================================================================')
    
    rules = []
    
    for section in sections:
        if "FILE:" in section:
            filename = section.split('\n')[1].strip()
            text = section
            
            print(f"--- Analyzing {filename} ---")
            
            # 1. Motion Deadlines & Notice
            # Look for 16 court days, 75 days, etc.
            # Regex for "X days" near "before hearing" or "service"
            
            # Common patterns:
            # "16 court days"
            # "75 days"
            # "CCP 1005"
            # "CCP 437c"
            
            # Extract citations
            citations = re.findall(r"(?:CCP|Code Civ\. Proc\.|§)\s*(\d+[a-z]?(?:\.\d+)?)", text)
            citations = sorted(list(set(citations)))
            print(f"Found Citations: {citations[:10]}...")

            # Extract Service Extensions
            # Mail: 5 days (CA), 10 (outside CA), 20 (outside US)
            # Fax/Email/Overnight: 2 days
            
            service_patterns = [
                (r"mail.*?(\d+)\s*(calendar)?\s*days", "Mail Extension"),
                (r"express.*?(\d+)\s*(calendar)?\s*days", "Express/Overnight Extension"),
                (r"fax.*?(\d+)\s*(calendar)?\s*days", "Fax Extension"),
                (r"electronic.*?(\d+)\s*(calendar)?\s*days", "Electronic Service Extension"),
            ]
            
            print("\n[Service Extensions Found]")
            for pat, name in service_patterns:
                matches = re.findall(pat, text, re.I)
                if matches:
                    print(f"  {name}: {matches[0][0]} days")

            # Extract Motion specific rules
            print("\n[Motion Deadlines]")
            # Look for "Summary Judgment" logic
            if "Summary Judgment" in text or "437c" in text:
                msj_notice = re.search(r"75\s*days", text)
                if msj_notice:
                    print("  MSJ Notice: 75 days (likely)")
            
            # Look for "16 court days"
            if "16 court days" in text:
                print("  General Motion Notice: 16 court days")

            # Opposition / Reply
            opp_match = re.search(r"opposition.*?(\d+)\s*(court|calendar)?\s*days", text, re.I)
            if opp_match:
                print(f"  Opposition: {opp_match.group(1)} {opp_match.group(2) or 'days'}")
                
            reply_match = re.search(r"reply.*?(\d+)\s*(court|calendar)?\s*days", text, re.I)
            if reply_match:
                print(f"  Reply: {reply_match.group(1)} {reply_match.group(2) or 'days'}")

            # Holidays
            if "12a" in text or "holiday" in text.lower():
                 print("\n[Holiday Rules]")
                 print("  References to CCP 12a or holidays found.")

            print("\n" + "-"*40 + "\n")

if __name__ == "__main__":
    extract_rules("extracted_text_word.txt")
