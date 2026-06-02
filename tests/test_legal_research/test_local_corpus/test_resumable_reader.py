"""The resumable HTTP reader must deliver a gap-free byte stream across a
dropped connection by reconnecting with a Range request."""
import io

from icharlotte_core.legal_research.local_corpus.build import _ResumableHTTPReader


class _FakeResp:
    def __init__(self, data: bytes, *, fail_after: int | None = None, content_length: int | None = None):
        self._buf = io.BytesIO(data)
        self._fail_after = fail_after   # raise ConnectionError after N bytes served
        self._served = 0
        self.headers = {"Content-Length": str(content_length)} if content_length is not None else {}

    def read(self, size=-1):
        if self._fail_after is not None and self._served >= self._fail_after:
            raise ConnectionResetError("simulated drop")
        chunk = self._buf.read(size)
        self._served += len(chunk)
        return chunk

    def close(self):
        pass


def test_resumes_across_dropped_connection():
    full = bytes(range(256)) * 40        # 10,240 bytes of known content
    calls = {"n": 0}

    def opener(url, headers, timeout):
        calls["n"] += 1
        start = 0
        if "Range" in headers:
            # bytes=START-
            start = int(headers["Range"].split("=")[1].split("-")[0])
        if calls["n"] == 1:
            # First connection: serve from 0 but die after 3000 bytes.
            return _FakeResp(full[start:], fail_after=3000, content_length=len(full))
        # Reconnect(s): serve the remainder from `start` cleanly.
        return _FakeResp(full[start:], content_length=len(full))

    reader = _ResumableHTTPReader("http://x/file", opener=opener, max_retries=5, backoff=0.0)
    got = b""
    while True:
        b = reader.read(1024)
        if not b:
            break
        got += b
    assert got == full                   # gap-free, no duplication, full content
    assert calls["n"] >= 2               # it actually reconnected
    assert reader.offset == len(full)


def test_clean_eof_without_drop():
    full = b"hello world" * 100
    def opener(url, headers, timeout):
        return _FakeResp(full, content_length=len(full))
    reader = _ResumableHTTPReader("http://x/f", opener=opener, backoff=0.0)
    assert reader.read(-1) == full
    assert reader.read(1024) == b""      # EOF
