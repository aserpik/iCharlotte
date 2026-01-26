import sqlite3
import os
from datetime import datetime
from .config import GEMINI_DATA_DIR


class TemplatesDatabase:
    """Database operations for templates and resources metadata."""

    def __init__(self):
        self.db_path = os.path.join(GEMINI_DATA_DIR, "master_cases.db")
        self.init_tables()

    def init_tables(self):
        """Initialize templates and resources tables."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Templates metadata table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                relative_path TEXT NOT NULL UNIQUE,
                category TEXT DEFAULT '',
                description TEXT DEFAULT '',
                created_date TEXT,
                modified_date TEXT
            )
        ''')

        # Template tags (many-to-many)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS template_tags (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                template_id INTEGER NOT NULL,
                tag TEXT NOT NULL,
                FOREIGN KEY(template_id) REFERENCES templates(id) ON DELETE CASCADE,
                UNIQUE(template_id, tag)
            )
        ''')

        # Resource tags
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS resource_tags (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                resource_path TEXT NOT NULL,
                tag TEXT NOT NULL,
                UNIQUE(resource_path, tag)
            )
        ''')

        # Placeholder mappings (global custom placeholder -> case variable mappings)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS placeholder_mappings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                custom_name TEXT NOT NULL UNIQUE,
                maps_to TEXT NOT NULL,
                created_date TEXT
            )
        ''')

        conn.commit()
        conn.close()

    # ===== Template Methods =====

    def upsert_template(self, relative_path, filename, category='', description=''):
        """Insert or update a template record."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        now = datetime.now().isoformat()

        cursor.execute('''
            INSERT INTO templates (filename, relative_path, category, description, created_date, modified_date)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(relative_path) DO UPDATE SET
                filename = excluded.filename,
                category = excluded.category,
                description = excluded.description,
                modified_date = excluded.modified_date
        ''', (filename, relative_path, category, description, now, now))

        template_id = cursor.lastrowid
        if template_id == 0:
            # Was an update, get the actual ID
            cursor.execute('SELECT id FROM templates WHERE relative_path = ?', (relative_path,))
            row = cursor.fetchone()
            template_id = row[0] if row else None

        conn.commit()
        conn.close()
        return template_id

    def get_template(self, template_id):
        """Get a template by ID."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('SELECT * FROM templates WHERE id = ?', (template_id,))
        row = cursor.fetchone()
        conn.close()

        return dict(row) if row else None

    def get_template_by_path(self, relative_path):
        """Get a template by its relative path."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('SELECT * FROM templates WHERE relative_path = ?', (relative_path,))
        row = cursor.fetchone()
        conn.close()

        return dict(row) if row else None

    def get_all_templates(self, category=None):
        """Get all templates, optionally filtered by category."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        if category:
            cursor.execute('SELECT * FROM templates WHERE category = ? ORDER BY filename', (category,))
        else:
            cursor.execute('SELECT * FROM templates ORDER BY filename')

        rows = cursor.fetchall()
        conn.close()

        return [dict(row) for row in rows]

    def delete_template(self, template_id):
        """Delete a template and its tags."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('DELETE FROM template_tags WHERE template_id = ?', (template_id,))
        cursor.execute('DELETE FROM templates WHERE id = ?', (template_id,))

        conn.commit()
        conn.close()

    def delete_template_by_path(self, relative_path):
        """Delete a template by path."""
        template = self.get_template_by_path(relative_path)
        if template:
            self.delete_template(template['id'])

    def update_template_category(self, template_id, category):
        """Update a template's category."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('UPDATE templates SET category = ?, modified_date = ? WHERE id = ?',
                      (category, datetime.now().isoformat(), template_id))

        conn.commit()
        conn.close()

    def get_all_categories(self):
        """Get list of all unique categories."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('SELECT DISTINCT category FROM templates WHERE category != "" ORDER BY category')
        rows = cursor.fetchall()
        conn.close()

        return [row[0] for row in rows]

    # ===== Template Tag Methods =====

    def add_template_tag(self, template_id, tag):
        """Add a tag to a template."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        try:
            cursor.execute('INSERT INTO template_tags (template_id, tag) VALUES (?, ?)',
                          (template_id, tag.strip().lower()))
            conn.commit()
        except sqlite3.IntegrityError:
            pass  # Tag already exists

        conn.close()

    def remove_template_tag(self, template_id, tag):
        """Remove a tag from a template."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('DELETE FROM template_tags WHERE template_id = ? AND tag = ?',
                      (template_id, tag.strip().lower()))

        conn.commit()
        conn.close()

    def get_template_tags(self, template_id):
        """Get all tags for a template."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('SELECT tag FROM template_tags WHERE template_id = ? ORDER BY tag',
                      (template_id,))
        rows = cursor.fetchall()
        conn.close()

        return [row[0] for row in rows]

    def get_templates_by_tag(self, tag):
        """Get all templates with a specific tag."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('''
            SELECT t.* FROM templates t
            JOIN template_tags tt ON t.id = tt.template_id
            WHERE tt.tag = ?
            ORDER BY t.filename
        ''', (tag.strip().lower(),))

        rows = cursor.fetchall()
        conn.close()

        return [dict(row) for row in rows]

    def get_all_template_tags(self):
        """Get list of all unique template tags."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('SELECT DISTINCT tag FROM template_tags ORDER BY tag')
        rows = cursor.fetchall()
        conn.close()

        return [row[0] for row in rows]

    # ===== Resource Tag Methods =====

    def add_resource_tag(self, resource_path, tag):
        """Add a tag to a resource."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        try:
            cursor.execute('INSERT INTO resource_tags (resource_path, tag) VALUES (?, ?)',
                          (resource_path, tag.strip().lower()))
            conn.commit()
        except sqlite3.IntegrityError:
            pass  # Tag already exists

        conn.close()

    def remove_resource_tag(self, resource_path, tag):
        """Remove a tag from a resource."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('DELETE FROM resource_tags WHERE resource_path = ? AND tag = ?',
                      (resource_path, tag.strip().lower()))

        conn.commit()
        conn.close()

    def get_resource_tags(self, resource_path):
        """Get all tags for a resource."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('SELECT tag FROM resource_tags WHERE resource_path = ? ORDER BY tag',
                      (resource_path,))
        rows = cursor.fetchall()
        conn.close()

        return [row[0] for row in rows]

    def get_resources_by_tag(self, tag):
        """Get all resource paths with a specific tag."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('SELECT resource_path FROM resource_tags WHERE tag = ? ORDER BY resource_path',
                      (tag.strip().lower(),))
        rows = cursor.fetchall()
        conn.close()

        return [row[0] for row in rows]

    def get_all_resource_tags(self):
        """Get list of all unique resource tags."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('SELECT DISTINCT tag FROM resource_tags ORDER BY tag')
        rows = cursor.fetchall()
        conn.close()

        return [row[0] for row in rows]

    def delete_resource_tags(self, resource_path):
        """Delete all tags for a resource."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('DELETE FROM resource_tags WHERE resource_path = ?', (resource_path,))

        conn.commit()
        conn.close()

    # ===== Placeholder Mapping Methods =====

    def add_placeholder_mapping(self, custom_name, maps_to):
        """Create or update a placeholder mapping."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            INSERT INTO placeholder_mappings (custom_name, maps_to, created_date)
            VALUES (?, ?, ?)
            ON CONFLICT(custom_name) DO UPDATE SET
                maps_to = excluded.maps_to
        ''', (custom_name.upper(), maps_to, datetime.now().isoformat()))

        conn.commit()
        conn.close()

    def get_placeholder_mapping(self, custom_name):
        """Get what a custom placeholder maps to."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('SELECT maps_to FROM placeholder_mappings WHERE custom_name = ?',
                      (custom_name.upper(),))
        row = cursor.fetchone()
        conn.close()

        return row[0] if row else None

    def get_all_placeholder_mappings(self):
        """Get all placeholder mappings."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('SELECT custom_name, maps_to FROM placeholder_mappings ORDER BY custom_name')
        rows = cursor.fetchall()
        conn.close()

        return {row['custom_name']: row['maps_to'] for row in rows}

    def delete_placeholder_mapping(self, custom_name):
        """Remove a placeholder mapping."""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('DELETE FROM placeholder_mappings WHERE custom_name = ?',
                      (custom_name.upper(),))

        conn.commit()
        conn.close()
