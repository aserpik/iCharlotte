"""
Download attachments containing 'carrier001' in filename from Outlook Sent Items,
excluding emails where the TO field contains jbordinwosk@bordinsemmer.com.
Auto-renames files to avoid conflicts.
"""

import win32com.client
import os
import sys

SAVE_DIR = r"C:\geminiterminal2\autodownloadreports"
EXCLUDE_TO_EMAIL = "jbordinwosk@bordinsemmer.com"
KEYWORD = "carrier001"


def get_unique_path(directory, filename):
    """Return a unique file path, appending _1, _2, etc. if file exists."""
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(directory, filename)
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{base}_{counter}{ext}")
        counter += 1
    return candidate


def to_addresses_contain_excluded(mail_item):
    """Check if the TO field contains the excluded email address."""
    try:
        # Use Recipients collection for accurate parsing
        for i in range(1, mail_item.Recipients.Count + 1):
            recip = mail_item.Recipients.Item(i)
            # Type 1 = olTo, Type 2 = olCC, Type 3 = olBCC
            if recip.Type == 1:  # TO recipient only
                addr = ""
                try:
                    # Try SMTP address first (most reliable)
                    addr = recip.AddressEntry.GetExchangeUser()
                    if addr:
                        addr = addr.PrimarySmtpAddress
                except Exception:
                    pass
                if not addr:
                    try:
                        addr = recip.AddressEntry.Address
                    except Exception:
                        addr = recip.Address
                if addr and EXCLUDE_TO_EMAIL.lower() in addr.lower():
                    return True
        return False
    except Exception as e:
        # Fallback: check the .To property string
        try:
            to_field = mail_item.To or ""
            return EXCLUDE_TO_EMAIL.lower() in to_field.lower()
        except Exception:
            return False


def main():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # olFolderSentMail = 5
    sent_folder = namespace.GetDefaultFolder(5)
    items = sent_folder.Items
    total_items = items.Count
    print(f"Found {total_items} items in Sent Items folder.")

    downloaded = 0
    skipped_to = 0
    skipped_no_match = 0
    checked = 0

    for i in range(1, total_items + 1):
        try:
            item = items.Item(i)
        except Exception:
            continue

        # Only process mail items (not meeting requests, etc.)
        try:
            if item.Class != 43:  # olMail = 43
                continue
        except Exception:
            continue

        checked += 1
        if checked % 200 == 0:
            print(f"  Checked {checked}/{total_items} emails... ({downloaded} attachments downloaded so far)")

        # Skip if TO field contains the excluded address
        if to_addresses_contain_excluded(item):
            skipped_to += 1
            continue

        # Check attachments
        try:
            att_count = item.Attachments.Count
            if att_count == 0:
                continue
        except Exception:
            continue

        found_match = False
        for j in range(1, att_count + 1):
            try:
                att = item.Attachments.Item(j)
                filename = att.FileName or ""
                if KEYWORD.lower() in filename.lower():
                    save_path = get_unique_path(SAVE_DIR, filename)
                    att.SaveAsFile(save_path)
                    subject = ""
                    try:
                        subject = item.Subject or "(no subject)"
                    except Exception:
                        subject = "(unknown)"
                    print(f"  Saved: {os.path.basename(save_path)}  (from: {subject})")
                    downloaded += 1
                    found_match = True
            except Exception as e:
                print(f"  Warning: could not save attachment: {e}")

        if not found_match:
            skipped_no_match += 1

    print(f"\nDone!")
    print(f"  Emails checked: {checked}")
    print(f"  Skipped (TO contains excluded addr): {skipped_to}")
    print(f"  Attachments downloaded: {downloaded}")
    print(f"  Saved to: {SAVE_DIR}")


if __name__ == "__main__":
    main()
