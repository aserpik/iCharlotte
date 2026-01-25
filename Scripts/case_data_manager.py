import os
import json
import logging
import datetime
from typing import Any, Dict, List, Optional

# Import Tagging Engine
try:
    from tagging_engine import TaggingEngine
except ImportError:
    from Scripts.tagging_engine import TaggingEngine

class CaseDataManager:
    def __init__(self, base_dir: str = None):
        """
        Args:
            base_dir: Directory where .json files are stored. 
                      Defaults to .gemini/case_data
        """
        if base_dir:
            self.data_dir = base_dir
        else:
            self.data_dir = os.path.join(os.getcwd(), ".gemini", "case_data")
            
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
            
        self.tagging_engine = TaggingEngine()

    def _get_file_path(self, file_num: str) -> str:
        return os.path.join(self.data_dir, f"{file_num}.json")

    def _load_json(self, file_num: str) -> Dict:
        path = self._get_file_path(file_num)
        if not os.path.exists(path):
            return {}
        try:
            with open(path, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError:
            print(f"Error decoding JSON for {file_num}")
            return {}

    def _save_json(self, file_num: str, data: Dict):
        path = self._get_file_path(file_num)
        with open(path, 'w') as f:
            json.dump(data, f, indent=4)

    def get_variable_obj(self, file_num: str, key: str) -> Optional[Dict]:
        """Returns the full object {value, tags, source, timestamp}."""
        data = self._load_json(file_num)
        val = data.get(key)
        
        if val is None:
            return None
            
        # If it's the old format (raw value), migrate it temporarily in memory
        if not isinstance(val, dict) or "value" not in val:
            return {
                "value": val,
                "tags": [],
                "source": "legacy",
                "timestamp": None
            }
        
        return val

    def get_value(self, file_num: str, key: str) -> Any:
        """Returns ONLY the raw value. Compatible with legacy scripts."""
        obj = self.get_variable_obj(file_num, key)
        if obj:
            return obj.get("value")
        return None

    def save_variable(self, file_num: str, key: str, value: Any, 
                      source: str = "agent", 
                      auto_tag: bool = True,
                      extra_tags: List[str] = None):
        """
        Saves a variable.
        If auto_tag is True, runs the TaggingEngine.
        Includes simple file locking to prevent race conditions.
        """
        import time
        
        # Determine Tags (Perform outside lock to avoid holding it during slow LLM calls)
        tags = []
        if extra_tags:
            tags.extend(extra_tags)
            
        if auto_tag:
            # Fetch known parties to help standardize names
            known_parties = []
            try:
                # We need to read from the file to get current parties.
                # Since we are about to lock and write, reading now is safe enough for hints.
                # Warning: recursive calls to get_value -> load_json might be slightly inefficient but acceptable.
                plaintiffs = self.get_value(file_num, "plaintiffs")
                if isinstance(plaintiffs, list):
                    known_parties.extend([str(p) for p in plaintiffs])
                elif isinstance(plaintiffs, str):
                    known_parties.append(plaintiffs)
                    
                defendants = self.get_value(file_num, "defendants")
                if isinstance(defendants, list):
                    known_parties.extend([str(d) for d in defendants])
                elif isinstance(defendants, str):
                    known_parties.append(defendants)
            except Exception:
                # If we fail to read parties (e.g. file doesn't exist yet), just proceed without them
                pass

            print(f"[CaseDataManager] Generating tags for '{key}'...")
            try:
                generated = self.tagging_engine.generate_tags(value, context_description=key, known_parties=known_parties)
                tags.extend(generated)
            except Exception as e:
                print(f"[CaseDataManager] Tagging failed for '{key}': {e}")
        
        # Deduplicate tags
        tags = list(set(tags))

        lock_path = self._get_file_path(file_num) + ".lock"
        
        # Acquire Lock
        max_retries = 100
        for i in range(max_retries):
            try:
                # Create lock file exclusively
                with open(lock_path, 'x') as f:
                    f.write(f"Locked by {os.getpid()}")
                break
            except FileExistsError:
                time.sleep(0.1)
                if i == max_retries - 1:
                    print(f"[CaseDataManager] Warning: Could not acquire lock for {file_num} after {max_retries} attempts. Proceeding anyway (risk of overwrite).")
        
        try:
            data = self._load_json(file_num)
            
            # Current time
            now = datetime.datetime.now().isoformat()
            
            # Construct Object
            new_obj = {
                "value": value,
                "tags": tags,
                "source": source,
                "timestamp": now
            }
            
            data[key] = new_obj
            self._save_json(file_num, data)
            print(f"[CaseDataManager] Saved '{key}' for {file_num}. Tags: {tags}")
            
        finally:
            # Release Lock
            if os.path.exists(lock_path):
                try:
                    os.remove(lock_path)
                except OSError:
                    pass

    def get_all_variables(self, file_num: str, flatten: bool = True) -> Dict:
        """
        Returns all variables.
        If flatten=True, returns {key: value} (legacy format).
        If flatten=False, returns {key: {value, tags...}}.
        """
        data = self._load_json(file_num)
        if not flatten:
            return data
            
        flat = {}
        for k, v in data.items():
            if isinstance(v, dict) and "value" in v:
                flat[k] = v["value"]
            else:
                flat[k] = v
        return flat

    def get_by_tag(self, file_num: str, tag: str) -> Dict[str, Any]:
        """Returns a dict of key:value for all variables containing the specified tag."""
        data = self._load_json(file_num)
        results = {}
        
        for k, v in data.items():
            if isinstance(v, dict) and "tags" in v:
                if tag in v["tags"]:
                    results[k] = v["value"]
                    
        return results
