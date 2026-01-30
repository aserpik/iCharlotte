import sqlite3
import os
from datetime import datetime
from .config import GEMINI_DATA_DIR

class MasterCaseDatabase:
    def __init__(self):
        self.db_path = os.path.join(GEMINI_DATA_DIR, "master_cases.db")
        self.init_db()

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Cases Table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cases (
                file_number TEXT PRIMARY KEY,
                plaintiff_last_name TEXT,
                next_hearing_date TEXT,
                trial_date TEXT,
                case_path TEXT
            )
        ''')
        
        # To-Do Table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS todos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_number TEXT,
                item TEXT,
                status TEXT DEFAULT 'pending',
                due_date TEXT,
                FOREIGN KEY(file_number) REFERENCES cases(file_number)
            )
        ''')

        # Check if 'color' column exists in 'todos', if not add it
        cursor.execute("PRAGMA table_info(todos)")
        columns = [info[1] for info in cursor.fetchall()]
        if 'color' not in columns:
            cursor.execute("ALTER TABLE todos ADD COLUMN color TEXT DEFAULT 'yellow'")
        
        # Add creation/assignment columns
        if 'created_date' not in columns:
            cursor.execute("ALTER TABLE todos ADD COLUMN created_date TEXT")
        if 'assigned_to' not in columns:
            cursor.execute("ALTER TABLE todos ADD COLUMN assigned_to TEXT")
        if 'assigned_date' not in columns:
            cursor.execute("ALTER TABLE todos ADD COLUMN assigned_date TEXT")

        # Check if 'assigned_attorney' column exists in 'cases', if not add it
        cursor.execute("PRAGMA table_info(cases)")
        case_columns = [info[1] for info in cursor.fetchall()]
        if 'assigned_attorney' not in case_columns:
            cursor.execute("ALTER TABLE cases ADD COLUMN assigned_attorney TEXT DEFAULT ''")
        if 'case_summary' not in case_columns:
            cursor.execute("ALTER TABLE cases ADD COLUMN case_summary TEXT DEFAULT ''")
        if 'last_report_text' not in case_columns:
            cursor.execute("ALTER TABLE cases ADD COLUMN last_report_text TEXT DEFAULT ''")
        if 'plaintiff_override' not in case_columns:
            cursor.execute("ALTER TABLE cases ADD COLUMN plaintiff_override INTEGER DEFAULT 0")
        if 'last_docket_download' not in case_columns:
            cursor.execute("ALTER TABLE cases ADD COLUMN last_docket_download TEXT")

        # History Table (Interim Updates / Status Reports)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_number TEXT,
                date TEXT,
                type TEXT,
                notes TEXT,
                FOREIGN KEY(file_number) REFERENCES cases(file_number)
            )
        ''')

        # Add email_entry_id column to history table (for linking to Outlook emails)
        cursor.execute("PRAGMA table_info(history)")
        history_columns = [info[1] for info in cursor.fetchall()]
        if 'email_entry_id' not in history_columns:
            cursor.execute("ALTER TABLE history ADD COLUMN email_entry_id TEXT")

        # Processed Emails Table (for Sent Items Monitor - prevents duplicate todos)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS processed_emails (
                entry_id TEXT PRIMARY KEY,
                file_number TEXT,
                processed_date TEXT,
                todo_text TEXT
            )
        ''')
        
        conn.commit()
        conn.close()

    def upsert_case(self, file_number, plaintiff_last_name, next_hearing_date, trial_date, case_path, plaintiff_override=0):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            # Use CASE statements to preserve plaintiff_last_name and plaintiff_override if already set to 1
            cursor.execute('''
                INSERT INTO cases (file_number, plaintiff_last_name, next_hearing_date, trial_date, case_path, plaintiff_override)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(file_number) DO UPDATE SET
                    plaintiff_last_name = CASE 
                        WHEN plaintiff_override = 1 AND excluded.plaintiff_override = 0 THEN plaintiff_last_name 
                        ELSE excluded.plaintiff_last_name 
                    END,
                    plaintiff_override = CASE 
                        WHEN plaintiff_override = 1 THEN 1 
                        ELSE excluded.plaintiff_override 
                    END,
                    next_hearing_date = CASE 
                        WHEN excluded.next_hearing_date != '' THEN excluded.next_hearing_date 
                        ELSE next_hearing_date 
                    END,
                    trial_date = CASE 
                        WHEN excluded.trial_date != '' THEN excluded.trial_date 
                        ELSE trial_date 
                    END,
                    case_path = CASE 
                        WHEN excluded.case_path != '' THEN excluded.case_path 
                        ELSE case_path 
                    END
            ''', (file_number, plaintiff_last_name, next_hearing_date, trial_date, case_path, plaintiff_override))
            conn.commit()
        finally:
            conn.close()

    def get_all_cases(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM cases ORDER BY file_number")
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def get_case(self, file_number):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM cases WHERE file_number = ?", (file_number,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def find_case_by_plaintiff(self, plaintiff_name):
        """Find a case by plaintiff name (case-insensitive partial match).

        Returns the first matching case dict or None if not found.
        """
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        # Use LIKE for case-insensitive partial match
        cursor.execute(
            "SELECT * FROM cases WHERE LOWER(plaintiff_last_name) LIKE LOWER(?) ORDER BY file_number",
            (f"%{plaintiff_name}%",)
        )
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def add_todo(self, file_number, item, due_date=None, color='yellow', created_date=None):
        if not created_date:
            created_date = datetime.now().strftime("%Y-%m-%d")

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("INSERT INTO todos (file_number, item, due_date, color, created_date, assigned_to, assigned_date) VALUES (?, ?, ?, ?, ?, ?, ?)", 
                           (file_number, item, due_date, color, created_date, "", ""))
            conn.commit()
        finally:
            conn.close()

    def get_todos(self, file_number):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM todos WHERE file_number = ? ORDER BY id DESC", (file_number,))
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def update_todo_status(self, todo_id, status):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE todos SET status = ? WHERE id = ?", (status, todo_id))
            conn.commit()
        finally:
            conn.close()

    def update_todo_color(self, todo_id, color):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE todos SET color = ? WHERE id = ?", (color, todo_id))
            conn.commit()
        finally:
            conn.close()

    def update_todo_assignment(self, todo_id, assigned_to, assigned_date):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE todos SET assigned_to = ?, assigned_date = ? WHERE id = ?", (assigned_to, assigned_date, todo_id))
            conn.commit()
        finally:
            conn.close()

    def update_assigned_attorney(self, file_number, initials):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE cases SET assigned_attorney = ? WHERE file_number = ?", (initials, file_number))
            conn.commit()
        finally:
            conn.close()

    def update_case_summary(self, file_number, summary):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE cases SET case_summary = ? WHERE file_number = ?", (summary, file_number))
            conn.commit()
        finally:
            conn.close()

    def update_plaintiff(self, file_number, name):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE cases SET plaintiff_last_name = ?, plaintiff_override = 1 WHERE file_number = ?", (name, file_number))
            conn.commit()
        finally:
            conn.close()

    def update_hearing_date(self, file_number, date_val):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE cases SET next_hearing_date = ? WHERE file_number = ?", (date_val, file_number))
            conn.commit()
        finally:
            conn.close()

    def update_trial_date(self, file_number, date_val):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE cases SET trial_date = ? WHERE file_number = ?", (date_val, file_number))
            conn.commit()
        finally:
            conn.close()

    def update_last_report_text(self, file_number, text_val):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE cases SET last_report_text = ? WHERE file_number = ?", (text_val, file_number))
            conn.commit()
        finally:
            conn.close()

    def update_last_docket_download(self, file_number, date_val):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE cases SET last_docket_download = ? WHERE file_number = ?", (date_val, file_number))
            conn.commit()
        finally:
            conn.close()
            
    def delete_todo(self, todo_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM todos WHERE id = ?", (todo_id,))
            conn.commit()
        finally:
            conn.close()

    def add_history(self, file_number, type_val, notes, date_val=None, email_entry_id=None):
        if not date_val:
            date_val = datetime.now().strftime("%Y-%m-%d")

        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("INSERT INTO history (file_number, date, type, notes, email_entry_id) VALUES (?, ?, ?, ?, ?)",
                           (file_number, date_val, type_val, notes, email_entry_id))
            conn.commit()
        finally:
            conn.close()

    def update_history_date(self, history_id, new_date):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE history SET date = ? WHERE id = ?", (new_date, history_id))
            conn.commit()
        finally:
            conn.close()

    def update_history_type(self, history_id, new_type):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("UPDATE history SET type = ? WHERE id = ?", (new_type, history_id))
            conn.commit()
        finally:
            conn.close()

    def delete_history(self, history_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM history WHERE id = ?", (history_id,))
            conn.commit()
        finally:
            conn.close()

    def get_history(self, file_number):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM history WHERE file_number = ? ORDER BY date DESC, id DESC", (file_number,))
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]

    def get_last_status_update_date(self, file_number):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT date FROM history WHERE file_number = ? AND type = 'Status Update' ORDER BY date DESC LIMIT 1", (file_number,))
            row = cursor.fetchone()
            return row[0] if row else None
        finally:
            conn.close()

    def delete_case(self, file_number):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            # Delete related data first to avoid orphans
            cursor.execute("DELETE FROM todos WHERE file_number = ?", (file_number,))
            cursor.execute("DELETE FROM history WHERE file_number = ?", (file_number,))
            cursor.execute("DELETE FROM cases WHERE file_number = ?", (file_number,))
            conn.commit()
        finally:
            conn.close()

    def clear_all_cases(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            # We only clear the cases list, but keeping todos/history might leave orphans.
            # Usually 'Clear List' implies a full reset of the list view.
            # If the user re-scans, they want a fresh start.
            # Ideally we should cascade delete or warn.
            # For now, let's just delete from 'cases'.
            # SQLite foreign keys are ON by default? No, usually OFF unless PRAGMA foreign_keys = ON.
            # So todos/history will likely remain but be orphaned.
            cursor.execute("DELETE FROM cases")
            conn.commit()
        finally:
            conn.close()

    def is_email_processed(self, entry_id):
        """Check if an email has already been processed (to prevent duplicate todos)."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT 1 FROM processed_emails WHERE entry_id = ?", (entry_id,))
            return cursor.fetchone() is not None
        finally:
            conn.close()

    def mark_email_processed(self, entry_id, file_number, todo_text):
        """Mark an email as processed after creating a todo from it."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute(
                "INSERT OR IGNORE INTO processed_emails (entry_id, file_number, processed_date, todo_text) VALUES (?, ?, ?, ?)",
                (entry_id, file_number, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), todo_text)
            )
            conn.commit()
        finally:
            conn.close()
