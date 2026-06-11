"""Backwards-compatible shim — the design system moved to icharlotte_core/ui/theme.py.

The wizard design system was promoted app-wide (it now styles the whole app via
``QApplication.setStyleSheet(theme.app_stylesheet())``). Existing wizard modules
import ``theme`` from this package, so re-export everything from the new home.

New code should import directly:
    from icharlotte_core.ui import theme
"""
from icharlotte_core.ui.theme import *  # noqa: F401,F403
from icharlotte_core.ui.theme import _asset_path, _btn_qss  # noqa: F401
