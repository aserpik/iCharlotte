"""Cyclic-GC critical-section guard.

`no_gc()` pauses Python's *generational (cyclic)* garbage collector for the
duration of a `with` block and restores the previous state on exit.

## Why this exists

iCharlotte has crashed repeatedly with a native access violation (0xc0000005)
preceded by a flood of `0x8001010e` (RPC_E_WRONG_THREAD). Root cause: pythoncom
`CDispatch` COM proxies (Outlook/Word) hold reference cycles. Python's cyclic
collector can run on *any* thread during allocation; if it finalizes one of those
proxies on a thread other than the STA apartment that created it, the cross-thread
COM Release fails with RPC_E_WRONG_THREAD and corrupts the heap. The next memory
access — typically deep in main-thread Qt/Shiboken code — dies with an access
violation. See memory/crash_20260604_com_wrong_thread.md and com_cyclic_gc_crash.md.

Reference-count deallocation is unaffected by `gc.disable()`; only the cyclic
collector is paused. So wrapping a *critical section* (a main-thread UI operation,
or a worker's COM block) in `no_gc()` guarantees the cyclic collector cannot fire
mid-section and finalize a COM proxy on the wrong apartment.

## Usage

Main-thread UI section (just prevent cyclic GC from firing here — do NOT force a
collect, which could itself finalize a leaked proxy at this point)::

    with no_gc(collect_on_exit=False):
        ...heavy Qt work (build a tab, populate the tree)...

Worker COM section (collect the proxies this unit created, on *this* thread,
before returning)::

    with no_gc(collect_on_exit=True):
        ...win32com work...
"""

import gc
from contextlib import contextmanager


@contextmanager
def no_gc(collect_on_exit: bool = True):
    """Disable the cyclic garbage collector inside the block.

    Args:
        collect_on_exit: When True, run a single ``gc.collect()`` on the current
            thread just before restoring GC state — useful inside a worker's COM
            section so cycles created there are reclaimed on the creating thread.
            Set False for main-thread UI sections, where forcing a collect could
            finalize a leaked cross-apartment proxy at exactly the wrong moment.

    The previous enabled/disabled state is always restored on exit (including when
    the body raises), so nested guards do not clobber an outer guard's state.
    """
    was_enabled = gc.isenabled()
    gc.disable()
    try:
        yield
    finally:
        if collect_on_exit:
            try:
                gc.collect()
            except Exception:
                pass
        if was_enabled:
            gc.enable()
