import sqlite3
import os
import datetime
from PyQt6.QtCore import QObject, pyqtSignal, QThread
import win32com.client
import pythoncom

from .config import GEMINI_DATA_DIR

class EmailDatabase:
    def __init__(self, case_number):
        self.case_number = case_number
        self.db_path = os.path.join(GEMINI_DATA_DIR, f"{case_number}_emails.db")
        self.init_db()

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Main email table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS emails (
                entry_id TEXT PRIMARY KEY,
                subject TEXT,
                sender TEXT,
                sender_email TEXT,
                to_recipients TEXT,
                cc_recipients TEXT,
                body_text TEXT,
                body_html TEXT,
                received_time TEXT,
                has_attachments BOOLEAN,
                conversation_id TEXT,
                folder_path TEXT
            )
        ''')
        
        # FTS5 Virtual Table for fast search
        # We trigger updates to this via application logic or triggers. 
        # Triggers are easier to keep in sync.
        cursor.execute('''
            CREATE VIRTUAL TABLE IF NOT EXISTS emails_fts USING fts5(
                subject, 
                sender, 
                body_text, 
                content='emails', 
                content_rowid='rowid'
            )
        ''')

        # Triggers to keep FTS in sync
        cursor.execute('''
            CREATE TRIGGER IF NOT EXISTS emails_ai AFTER INSERT ON emails BEGIN
                INSERT INTO emails_fts(rowid, subject, sender, body_text) 
                VALUES (new.rowid, new.subject, new.sender, new.body_text);
            END;
        ''')
        cursor.execute('''
            CREATE TRIGGER IF NOT EXISTS emails_ad AFTER DELETE ON emails BEGIN
                INSERT INTO emails_fts(emails_fts, rowid, subject, sender, body_text) 
                VALUES('delete', old.rowid, old.subject, old.sender, old.body_text);
            END;
        ''')
        cursor.execute('''
            CREATE TRIGGER IF NOT EXISTS emails_au AFTER UPDATE ON emails BEGIN
                INSERT INTO emails_fts(emails_fts, rowid, subject, sender, body_text) 
                VALUES('delete', old.rowid, old.subject, old.sender, old.body_text);
                INSERT INTO emails_fts(rowid, subject, sender, body_text) 
                VALUES (new.rowid, new.subject, new.sender, new.body_text);
            END;
        ''')

        # Metadata table for sync state
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS metadata (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        ''')
        
        conn.commit()
        conn.close()

    def get_last_sync_time(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT value FROM metadata WHERE key='last_sync_time'")
            row = cursor.fetchone()
            return row[0] if row else None
        except:
            return None
        finally:
            conn.close()

    def set_last_sync_time(self, timestamp_str):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("INSERT OR REPLACE INTO metadata (key, value) VALUES ('last_sync_time', ?)", (timestamp_str,))
            conn.commit()
        finally:
            conn.close()

    def upsert_email(self, email_data):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO emails (
                    entry_id, subject, sender, sender_email, to_recipients, cc_recipients, 
                    body_text, body_html, received_time, has_attachments, conversation_id, folder_path
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(entry_id) DO UPDATE SET
                    subject=excluded.subject,
                    body_text=excluded.body_text,
                    body_html=excluded.body_html,
                    has_attachments=excluded.has_attachments,
                    folder_path=excluded.folder_path
            ''', (
                email_data['entry_id'],
                email_data['subject'],
                email_data['sender'],
                email_data['sender_email'],
                email_data['to'],
                email_data['cc'],
                email_data['body_text'],
                email_data['body_html'],
                email_data['received_time'],
                email_data['has_attachments'],
                email_data['conversation_id'],
                email_data['folder_path']
            ))
            conn.commit()
        except Exception as e:
            print(f"Error upserting email: {e}")
        finally:
            conn.close()

    def search_emails(self, query=None, limit=2000):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        if query:
            # FTS Search
            # Order by rank is standard for FTS
            sql = f'''
                SELECT e.* FROM emails e
                JOIN emails_fts f ON e.rowid = f.rowid
                WHERE emails_fts MATCH ?
                ORDER BY rank
                LIMIT ?
            '''
            # FTS5 query syntax: replace special chars or just pass raw if simple
            # Simple sanitization to prevent syntax errors in match query
            safe_query = query.replace('"', '""')
            cursor.execute(sql, (safe_query, limit))
        else:
            # Recent emails
            sql = '''
                SELECT * FROM emails 
                ORDER BY received_time DESC 
                LIMIT ?
            '''
            cursor.execute(sql, (limit,))
            
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def get_chronological_emails(self):
        """
        Retrieves all emails ordered by date.
        """
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # Select all fields to ensure we have data for display (html, etc)
        sql = '''
            SELECT * 
            FROM emails 
            ORDER BY received_time ASC
        '''
        cursor.execute(sql)
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]

class EmailSyncWorker(QThread):
    progress = pyqtSignal(str) # Status messages
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, case_number, full_sync=False):
        super().__init__()
        self.case_number = case_number
        self.full_sync = full_sync
        self.stop_requested = False

    def run(self):
        try:
            from .utils import log_event
            import time
            pythoncom.CoInitialize()
            self.db = EmailDatabase(self.case_number)
            
            # Retrieve last sync time
            last_sync_str = self.db.get_last_sync_time()
            
            if self.full_sync or not last_sync_str:
                sync_cutoff = '1900-01-01 00:00:00'
                log_event(f"EmailSync: Full sync initiated (from {sync_cutoff})")
            else:
                # Use last sync time
                sync_cutoff = last_sync_str
                log_event(f"EmailSync: Incremental sync from {sync_cutoff}")
            
            # Capture current time for next sync (using UTC or consistent local time)
            # Outlook timestamps are often local or UTC depending on property. 
            # We'll use local time as the marker for the next run.
            current_sync_start = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            log_event(f"EmailSync: Starting AdvancedSearch (Server-Side) for {self.case_number}")
            outlook = win32com.client.Dispatch("Outlook.Application")
            mapi = outlook.GetNamespace("MAPI")
            
            self.total_processed = 0
            processed_entry_ids = set()
            batch_data = []

            # --- Helper: Batch Saver ---
            def commit_batch():
                if not batch_data: return
                try:
                    conn = sqlite3.connect(self.db.db_path)
                    cursor = conn.cursor()
                    cursor.executemany('''
                        INSERT INTO emails (
                            entry_id, subject, sender, sender_email, to_recipients, cc_recipients, 
                            body_text, body_html, received_time, has_attachments, conversation_id, folder_path
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(entry_id) DO UPDATE SET
                            subject=excluded.subject,
                            body_text=excluded.body_text,
                            body_html=excluded.body_html,
                            has_attachments=excluded.has_attachments,
                            folder_path=excluded.folder_path
                    ''', batch_data)
                    conn.commit()
                    conn.close()
                    batch_data.clear()
                except Exception as e:
                    log_event(f"EmailSync: Batch save error: {e}", "error")

            def buffer_item(item, folder_path):
                try:
                    try:
                        sender_name = item.SenderName
                        sender_email = item.SenderEmailAddress
                    except:
                        sender_name = "Unknown"
                        sender_email = "unknown"

                    row = (
                        item.EntryID,
                        item.Subject,
                        sender_name,
                        sender_email,
                        item.To,
                        item.CC,
                        item.Body,
                        item.HTMLBody if hasattr(item, "HTMLBody") else "", 
                        str(item.ReceivedTime),
                        item.Attachments.Count > 0,
                        item.ConversationID if hasattr(item, "ConversationID") else "",
                        folder_path
                    )
                    batch_data.append(row)
                    if len(batch_data) >= 50:
                        commit_batch()
                except: pass
            # ---------------------------

            # 1. SEARCH SPECIFIC FOLDERS (Optimized Direct Access)
            # Strategy: Look directly for "CASES" at the root of every store.
            target_folders = []
            
            def normalize_for_match(s):
                return "".join(c for c in s if c.isalnum()).lower()
            
            normalized_case_num = normalize_for_match(self.case_number)
            self.progress.emit(f"Locating case folders (matching '{normalized_case_num}')...")
            
            for store in mapi.Stores:
                try:
                    root = store.GetRootFolder()
                    cases_folder = None
                    
                    # Case-insensitive check for CASES folder
                    for f in root.Folders:
                        if f.Name.lower() == "cases":
                            cases_folder = f
                            break
                            
                    if cases_folder:
                        log_event(f"EmailSync: Found CASES in {store.DisplayName}")
                        
                        # Iterate ONLY the CASES folder looking for the file number
                        count_checked = 0
                        for sub in cases_folder.Folders:
                            count_checked += 1
                            # Normalize folder name comparison
                            if normalized_case_num in normalize_for_match(sub.Name):
                                log_event(f"EmailSync: Matched case folder: {sub.Name}")
                                target_folders.append(sub)
                            else:
                                # excessive logging, but helpful for debug if needed, maybe commented out usually
                                # log_event(f"EmailSync: Ignored: {sub.Name}")
                                pass
                        log_event(f"EmailSync: Checked {count_checked} subfolders in {store.DisplayName}/CASES")
                    else:
                         log_event(f"EmailSync: No CASES folder in {store.DisplayName}")

                except Exception as e:
                    log_event(f"EmailSync: Error scanning store {store.DisplayName}: {e}", "warning")

            log_event(f"EmailSync: Found {len(target_folders)} target folders.")
            
            # Define Property Tags (Schemas) for GetTable
            # We use these to fetch raw data without creating full Item objects (Massive Speedup)
            PR_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"
            PR_SUBJECT = "http://schemas.microsoft.com/mapi/proptag/0x0037001F"
            PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
            PR_SENDER_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F" 
            PR_DISPLAY_TO = "http://schemas.microsoft.com/mapi/proptag/0x0E04001F"
            PR_DISPLAY_CC = "http://schemas.microsoft.com/mapi/proptag/0x0E03001F"
            PR_BODY = "http://schemas.microsoft.com/mapi/proptag/0x1000001F"
            PR_HTML_BODY = "http://schemas.microsoft.com/mapi/proptag/0x10130102" # Binary
            PR_CLIENT_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
            PR_MESSAGE_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
            PR_HAS_ATTACH = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"
            PR_CONVERSATION_ID = "http://schemas.microsoft.com/mapi/proptag/0x30130102"

            cols = [
                PR_ENTRYID, PR_SUBJECT, PR_SENDER_NAME, PR_SENDER_EMAIL, 
                PR_DISPLAY_TO, PR_DISPLAY_CC, PR_BODY, PR_CLIENT_SUBMIT_TIME, 
                PR_MESSAGE_DELIVERY_TIME,
                PR_HAS_ATTACH, PR_CONVERSATION_ID
            ]

            def process_folder_fast(folder, restriction=None):
                if self.stop_requested: return
                try:
                    # Robust iteration using folder.Items (GetTable is flaky for Body/HTML)
                    items = folder.Items
                    if restriction:
                        try:
                            items = items.Restrict(restriction)
                        except: pass # Fallback to all items if restrict fails
                    
                    # Sort not strictly necessary but nice for logs? No, skip for speed.
                    
                    count = 0
                    for item in items:
                        if self.stop_requested: break
                        try:
                            # Verify message class if needed, or just try-except keys
                            # We want IPM.Note mostly, but others might be relevant.
                            
                            eid = item.EntryID
                            if eid in processed_entry_ids:
                                continue
                            
                            buffer_item(item, folder.FolderPath)
                            processed_entry_ids.add(eid)
                            self.total_processed += 1
                            count += 1
                            
                            if count % 20 == 0:
                                self.progress.emit(f"Synced {self.total_processed} items...")
                                commit_batch()
                                
                        except Exception:
                            # Skip items that fail property access (e.g. strict permission or corrupted)
                            continue
                            
                    log_event(f"EmailSync: Scanned {count} items in {folder.Name}")

                except Exception as e:
                    log_event(f"EmailSync: Failed to process folder {folder.Name}: {e}", "error")

            # Sync Targets
            for folder in target_folders:
                if self.stop_requested: break
                
                # Get all subfolders
                subs = [folder]
                def recurse(f):
                    try:
                        for s in f.Folders: 
                            subs.append(s)
                            recurse(s)
                    except: pass
                recurse(folder)
                
                for f in subs:
                    if self.stop_requested: break
                    self.progress.emit(f"Scanning: {f.Name}...")
                    # Process with Fast Table
                    # For specific case folders, scan ALL items to catch moved/older emails
                    process_folder_fast(f, None)

            # 2. GLOBAL SEARCH (Inbox & Sent Items in Primary Account ONLY)
            if not self.stop_requested:
                self.progress.emit("Scanning Inbox & Sent Items...")
                
                search_roots = []
                try: search_roots.append(mapi.GetDefaultFolder(6)) # Inbox
                except: pass
                try: search_roots.append(mapi.GetDefaultFolder(5)) # Sent Items
                except: pass

                # Simple, direct restriction as requested
                # DASL syntax for exact substring match in subject
                restriction = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{self.case_number}%'"
                log_event(f"EmailSync: Global Search (Primary Account) using: {restriction}")

                for folder in search_roots:
                    if self.stop_requested: break
                    
                    self.progress.emit(f"Searching {folder.Name}...")
                    try:
                        # Restrict is often more thorough than GetTable for specific filters
                        items = folder.Items.Restrict(restriction)
                        count = 0
                        
                        for item in items:
                            if self.stop_requested: break
                            
                            # Skip if already processed in this run (e.g. from CASES folder)
                            try:
                                eid = item.EntryID
                                if eid in processed_entry_ids:
                                    continue
                                    
                                buffer_item(item, folder.FolderPath)
                                processed_entry_ids.add(eid)
                                count += 1
                                self.total_processed += 1
                                
                                if count % 20 == 0:
                                    self.progress.emit(f"Found {count} in {folder.Name}...")
                                    commit_batch()
                            except: continue
                            
                        log_event(f"EmailSync: Found {count} items in {folder.Name}")
                    except Exception as e:
                        log_event(f"EmailSync: Error searching {folder.Name}: {e}", "warning")

            commit_batch()

            commit_batch()
            
            # Update last sync time on success
            self.db.set_last_sync_time(current_sync_start)
            
            final_msg = f"EmailSync: Complete. {self.total_processed} items found."
            log_event(final_msg)
            self.progress.emit(final_msg.replace("EmailSync: ", ""))
            self.finished.emit()

        except Exception as e:
            log_event(f"EmailSync: Error: {e}", "error")
            self.error.emit(str(e))
        finally:
            pythoncom.CoUninitialize()

    def save_item(self, item, folder_path):
        try:
            try:
                sender_name = item.SenderName
                sender_email = item.SenderEmailAddress
            except:
                sender_name = "Unknown"
                sender_email = "unknown"

            email_data = {
                'entry_id': item.EntryID,
                'subject': item.Subject,
                'sender': sender_name,
                'sender_email': sender_email,
                'to': item.To,
                'cc': item.CC,
                'body_text': item.Body,
                'body_html': item.HTMLBody if hasattr(item, "HTMLBody") else "", 
                'received_time': str(item.ReceivedTime),
                'has_attachments': item.Attachments.Count > 0,
                'conversation_id': item.ConversationID if hasattr(item, "ConversationID") else "",
                'folder_path': folder_path
            }
            self.db.upsert_email(email_data)
        except Exception as e:
            print(f"Error saving item: {e}")

    def stop(self):
        self.stop_requested = True
