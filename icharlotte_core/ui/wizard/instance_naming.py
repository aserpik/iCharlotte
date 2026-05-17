"""Compute the next instance suffix for a duplicated task tab title."""
import re
from typing import Iterable


def next_instance_suffix(base_title: str, existing_titles: Iterable[str]) -> str:
    """Return the suffix to append to `base_title` for a new task tab.

    - If no existing tab uses `base_title` (with or without suffix), returns "".
    - Otherwise returns "(N)" where N is the lowest positive integer >= 2 that
      isn't already taken by a tab titled `base_title (N)`.

    Examples:
      base_title="Summarize Documents", existing=[] -> ""
      existing=["Summarize Documents"] -> "(2)"
      existing=["Summarize Documents", "Summarize Documents (3)"] -> "(2)"
    """
    existing = list(existing_titles)
    pattern = re.compile(rf"^{re.escape(base_title)}(?: \((\d+)\))?$")
    used_ns: set[int] = set()
    has_base = False
    for t in existing:
        m = pattern.match(t)
        if not m:
            continue
        num_str = m.group(1)
        if num_str is None:
            has_base = True
            used_ns.add(1)
        else:
            used_ns.add(int(num_str))
    if not has_base:
        return ""
    n = 2
    while n in used_ns:
        n += 1
    return f"({n})"
