"""Tests for icharlotte_core.gc_guard.no_gc — the cyclic-GC critical-section guard.

This guard prevents Python's generational (cyclic) garbage collector from firing
on the current thread inside a critical section. The motivating bug: cyclic GC can
run on ANY thread during allocation and finalize a pythoncom CDispatch COM proxy on
the wrong COM apartment, raising RPC_E_WRONG_THREAD (0x8001010e) and corrupting the
heap -> native access violation. See memory/crash_20260604_com_wrong_thread.md.
"""
import gc

import pytest

from icharlotte_core.gc_guard import no_gc


def test_no_gc_disables_cyclic_gc_inside_context():
    gc.enable()
    assert gc.isenabled()
    with no_gc():
        assert not gc.isenabled(), "cyclic GC must be disabled inside the guard"
    assert gc.isenabled(), "GC must be re-enabled after the guard (prior state)"


def test_no_gc_restores_prior_disabled_state():
    # If GC was already disabled going in, exiting must NOT re-enable it
    # (so nested guards don't clobber an outer guard's state).
    gc.disable()
    try:
        assert not gc.isenabled()
        with no_gc():
            assert not gc.isenabled()
        assert not gc.isenabled(), "must stay disabled when prior state was disabled"
    finally:
        gc.enable()  # restore sane default for the rest of the suite


def test_no_gc_restores_state_even_on_exception():
    gc.enable()
    with pytest.raises(ValueError):
        with no_gc():
            assert not gc.isenabled()
            raise ValueError("boom")
    assert gc.isenabled(), "GC must be restored even when the body raises"


def test_no_gc_collect_on_exit_does_not_error_and_restores():
    gc.enable()
    with no_gc(collect_on_exit=True):
        pass
    assert gc.isenabled()


def test_no_gc_collect_on_exit_false_skips_collect_but_restores():
    gc.enable()
    with no_gc(collect_on_exit=False):
        assert not gc.isenabled()
    assert gc.isenabled()
