# Separator Sensitivity Control — Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a 3-position sensitivity toggle (Broad/Default/Fine) to IndexTab that controls how aggressively the document separator groups or splits documents.

**Architecture:** The sensitivity value is a CLI arg (`--sensitivity 1|2|3`) threaded from IndexTab UI through `iCharlotte.run_separator_path()` to `Scripts/separate.py`, where it modifies the LLM prompt rules in `analyze_headers_chunk()`. UI provides a slider + re-analyze button above the document table.

**Tech Stack:** Python, PyQt6, argparse, Gemini API (existing)

**Spec:** `docs/superpowers/specs/2026-03-18-separator-sensitivity-design.md`

---

### Task 1: Add `--sensitivity` CLI arg and thread through call chain in `separate.py`

**Files:**
- Modify: `Scripts/separate.py:724-765` (main/argparse), `:406-447` (run_analysis), `:293-346` (analyze_headers), `:230-291` (analyze_headers_chunk)

- [ ] **Step 1: Add `--sensitivity` argument to argparse and update structured_args check**

In `main()` at line 729, after the `--headless` arg, add:

```python
parser.add_argument("--sensitivity", type=int, choices=[1, 2, 3], default=2,
                    help="Document separation sensitivity: 1=Broad, 2=Default, 3=Fine")
```

Also update the `structured_args` check on line 736 to include `--sensitivity`:
```python
structured_args = "--interactive" in sys.argv or "--headless" in sys.argv or "--original-pdf" in sys.argv or "--sensitivity" in sys.argv
```

Then pass it through in the structured args branch (~line 747):
```python
run_analysis(path, headless=args.headless, sensitivity=args.sensitivity)
```

- [ ] **Step 2: Thread sensitivity through `run_analysis` → `analyze_headers` → `analyze_headers_chunk`**

Update function signatures:

```python
def run_analysis(pdf_path, headless=False, sensitivity=2):
    # ... existing code ...
    docs = analyze_headers(headers, sensitivity)
    # ... rest unchanged ...
```

```python
def analyze_headers(headers, sensitivity=2):
    # ... existing code ...
    logger.info(f"Analyzing {total_pages} pages in chunks of {chunk_size}... (sensitivity={sensitivity})")
    # In the loop, pass sensitivity:
    chunk_docs = analyze_headers_chunk(chunk, start_page, next_id, prev_context, sensitivity)
    # ... rest unchanged ...
```

```python
def analyze_headers_chunk(headers_subset, start_page_num, next_id, prev_doc_context=None, sensitivity=2):
    # ... (prompt changes in next step) ...
```

- [ ] **Step 3: Modify prompt rules based on sensitivity**

In `analyze_headers_chunk`, replace the hardcoded rules string (lines 249-254) with sensitivity-conditional logic:

```python
    # Build rules based on sensitivity
    if sensitivity == 1:  # Broad
        rules = (
            "Rules:\n"
            "1. Return ONLY the list, one document per line.\n"
            "2. Do not use markdown.\n"
            "3. If a document continues to the end of this batch, set EndPage to the last page of this batch.\n"
            "4. Provide a detailed and descriptive title for each document.\n"
            "5. If you identify an Insurance Policy, group all its related parts (Declarations, Endorsements, Conditions, Exclusions, etc.) into a SINGLE document entry. Do not split it.\n"
            "6. Group related documents together liberally. For example, a motion and its exhibits should be ONE entry. A letter and its attachments should be ONE entry. Only create separate entries for clearly distinct, unrelated documents.\n"
        )
    elif sensitivity == 3:  # Fine
        rules = (
            "Rules:\n"
            "1. Return ONLY the list, one document per line.\n"
            "2. Do not use markdown.\n"
            "3. If a document continues to the end of this batch, set EndPage to the last page of this batch.\n"
            "4. Provide a detailed and descriptive title for each document.\n"
            "5. Be aggressive about identifying separate documents. Each exhibit, attachment, declaration, addendum, or sub-document should be its own entry. When in doubt, split rather than group.\n"
        )
    else:  # Default (2)
        rules = (
            "Rules:\n"
            "1. Return ONLY the list, one document per line.\n"
            "2. Do not use markdown.\n"
            "3. If a document continues to the end of this batch, set EndPage to the last page of this batch.\n"
            "4. Provide a detailed and descriptive title for each document.\n"
            "5. If you identify an Insurance Policy, group all its related parts (Declarations, Endorsements, Conditions, Exclusions, etc.) into a SINGLE document entry. Do not split it.\n"
        )
```

Then update the prompt construction to use `rules` variable instead of inline rules:

```python
    prompt = (
        "I am processing a large PDF in batches. This batch contains headers from pages "
        f"{start_page_num} to {start_page_num + len(headers_subset) - 1}.\n"
        "Identify distinct legal or administrative documents.\n"
        "Format: ID|Title|Date|StartPage|EndPage\n"
        f"{context_instruction}\n"
        f"{rules}"
        "Example:\n"
        "1|Plaintiff's Complaint|2023-01-01|1|5\n"
        "2|Exhibit A|2023-01-02|6|10\n"
        "\nHEADERS:\n" + "\n".join(headers_subset)
    )
```

- [ ] **Step 4: Verify script runs standalone**

Run: `python Scripts/separate.py --help`
Expected: Shows `--sensitivity` in help output with choices [1, 2, 3] and default 2.

- [ ] **Step 5: Commit**

```bash
git add Scripts/separate.py
git commit -m "feat(separator): add --sensitivity CLI arg for document grouping control"
```

---

### Task 2: Add sensitivity parameter to `run_separator_path` in `iCharlotte.py`

**Files:**
- Modify: `iCharlotte.py:1512-1559` (run_separator_path)

- [ ] **Step 1: Add sensitivity parameter to method signature**

Change line 1512:
```python
def run_separator_path(self, path, sensitivity=2):
```

- [ ] **Step 2: Pass sensitivity to CLI args**

Change line 1520 from:
```python
args = [script_path, "--headless", path]
```
to:
```python
args = [script_path, "--headless", "--sensitivity", str(sensitivity), path]
```

- [ ] **Step 3: Re-enable IndexTab controls on completion (success or failure)**

In the `on_finished` callback inside `run_separator_path` (~line 1532), add re-enable logic that runs regardless of success/failure. Add after `self.cleanup_runner(runner)` (line 1556):

```python
            # Re-enable sensitivity controls even on failure
            if hasattr(self, 'index_tab') and hasattr(self.index_tab, 'reanalyze_btn'):
                self.index_tab.reanalyze_btn.setEnabled(True)
                self.index_tab.sensitivity_slider.setEnabled(True)
```

- [ ] **Step 4: Commit**

```bash
git add iCharlotte.py
git commit -m "feat(separator): pass sensitivity param through run_separator_path"
```

---

### Task 3: Add sensitivity slider and re-analyze button to IndexTab UI

**Files:**
- Modify: `icharlotte_core/ui/tabs.py:1806-1960` (IndexTab.setup_ui)

- [ ] **Step 1: Add `QSlider` import and sensitivity controls**

First, add `QSlider` to the PyQt6/PySide6 imports at the top of `tabs.py`. Find the existing widget import line and add `QSlider` to it.

Then, after `middle_layout.addLayout(middle_header)` (line 1875) and before `self.doc_table = QTableWidget()` (line 1877), insert:

```python
        # Sensitivity slider + Re-analyze
        sensitivity_layout = QHBoxLayout()
        sensitivity_layout.setContentsMargins(4, 2, 4, 2)

        sensitivity_label = QLabel("Separation:")
        sensitivity_layout.addWidget(sensitivity_label)

        broad_label = QLabel("Broad")
        broad_label.setStyleSheet("color: #666; font-size: 11px;")
        sensitivity_layout.addWidget(broad_label)

        self.sensitivity_slider = QSlider(Qt.Orientation.Horizontal)
        self.sensitivity_slider.setMinimum(1)
        self.sensitivity_slider.setMaximum(3)
        self.sensitivity_slider.setValue(2)
        self.sensitivity_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.sensitivity_slider.setTickInterval(1)
        self.sensitivity_slider.setPageStep(1)
        self.sensitivity_slider.setFixedWidth(100)
        sensitivity_layout.addWidget(self.sensitivity_slider)

        fine_label = QLabel("Fine")
        fine_label.setStyleSheet("color: #666; font-size: 11px;")
        sensitivity_layout.addWidget(fine_label)

        self.reanalyze_btn = QPushButton("Re-analyze")
        self.reanalyze_btn.setToolTip("Re-run document separation with the selected sensitivity level")
        self.reanalyze_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 6px 12px;")
        self.reanalyze_btn.clicked.connect(self.on_reanalyze_clicked)
        sensitivity_layout.addWidget(self.reanalyze_btn)

        sensitivity_layout.addStretch()
        middle_layout.addLayout(sensitivity_layout)
```

- [ ] **Step 2: Add the `on_reanalyze_clicked` handler method to IndexTab**

Add this method to the IndexTab class:

```python
    def on_reanalyze_clicked(self):
        if not self.current_pdf_path:
            return
        sensitivity = self.sensitivity_slider.value()
        # Disable controls while running
        self.reanalyze_btn.setEnabled(False)
        self.sensitivity_slider.setEnabled(False)
        main_window = self.window()
        if hasattr(main_window, 'run_separator_path'):
            main_window.run_separator_path(self.current_pdf_path, sensitivity=sensitivity)
```

- [ ] **Step 3: Re-enable controls when analysis finishes**

In `add_pdf()` method (line 2257), add re-enable logic at the end:

```python
    def add_pdf(self, path, docs):
        self.index_data[path] = docs
        self.save_data()

        # Add to list if not present
        items = self.pdf_list.findItems(path, Qt.MatchFlag.MatchExactly)
        if not items:
            item = QListWidgetItem(path)
            item.setIcon(self.icon_provider.icon(QFileInfo(path)))
            self.pdf_list.addItem(item)
            self.pdf_list.setCurrentRow(self.pdf_list.count() - 1)
        else:
            self.pdf_list.setCurrentItem(items[0])
            self.on_pdf_selected(items[0], None)

        # Re-enable sensitivity controls
        if hasattr(self, 'reanalyze_btn'):
            self.reanalyze_btn.setEnabled(True)
            self.sensitivity_slider.setEnabled(True)
```

- [ ] **Step 4: Verify the app launches without errors**

Run: `python iCharlotte.py`
Expected: App opens, IndexTab shows sensitivity slider + Re-analyze button above the document table.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "feat(separator): add sensitivity slider and re-analyze button to IndexTab"
```

---

### Task 4: Manual end-to-end test

- [ ] **Step 1: Test default behavior**

1. Open iCharlotte, load a case with a multi-document PDF
2. Run the separator via file tree checkbox (existing flow)
3. Verify results appear in IndexTab — should be identical to previous behavior

- [ ] **Step 2: Test re-analyze with Fine sensitivity**

1. In IndexTab, select the same PDF
2. Move slider to "Fine" (rightmost position)
3. Click "Re-analyze"
4. Verify: more documents appear in the table than the default run
5. Verify: slider and button are disabled during analysis, re-enabled after

- [ ] **Step 3: Test re-analyze with Broad sensitivity**

1. Move slider to "Broad" (leftmost position)
2. Click "Re-analyze"
3. Verify: fewer documents appear in the table than the default run

- [ ] **Step 4: Commit final state (if any fixes were needed)**

```bash
git add Scripts/separate.py iCharlotte.py icharlotte_core/ui/tabs.py
git commit -m "fix(separator): address issues found during e2e testing"
```
