# Extramarks Dynamic Curriculum Creator — Complete Technical Architecture & Working Guide

Python + Playwright + OpenPyXL | Excel-driven curriculum automation | Architecture, execution flow, navigation, creation, verification and recovery

## 1. What exactly have we built?

The Dynamic Curriculum Creator is a browser-automation system designed to build a complete academic syllabus hierarchy inside an existing Board on the Extramarks platform.

Previously, your scripts could perform two separate operations:

- Legacy Code 1 — Content Population: Navigate an existing Board → Class → Subject → Chapter → Topic hierarchy and populate content into the selected Topic.
- Legacy Code 2 — Syllabus Creation: Create multiple Topics with their descriptions and three content groups under a hierarchy already opened in the portal.

The new package combines these capabilities to create an entire hierarchy automatically.

Core principle:

> Read the curriculum structure from Excel → Identify the correct parent → Check whether the child exists → Create if missing → Verify → Update progress → Continue.

### High-level system architecture

Curriculum Excel

Hierarchy, level type, description, groups, status

Excel Reader & Hierarchy Planner

Validation → Normalization → Parent-first sorting

Playwright Automation Engine

Navigator

Code 1 logic

Level Creator

Code 2 logic

Group Creator

Legacy group logic

Verification

Card detection

Extramarks Syllabus Portal

Create or reuse hierarchy levels

Status Tracking & Error Diagnostics

Workbook updates · Logs · Screenshots · JSON diagnostics

The system currently creates syllabus structure and group configurations. It does not create new Boards, upload PDFs, or populate actual learning content into those groups.

## 2. Package structure and technologies

The original ZIP package contains four files. The authentication file is supplied separately.

```
Curriculum_Creator/
│
├── curriculum_creator.py
│   ├── Configuration and constants
│   ├── Node data model
│   ├── ExcelState
│   ├── Portal
│   │   ├── Navigation methods
│   │   ├── Dropdown methods
│   │   ├── Level creation
│   │   ├── Group creation
│   │   └── Save verification
│   ├── Command-line argument parser
│   └── Main execution controller
│
├── curriculum_template.xlsx
├── README_curriculum_creator.md
├── requirements.txt
│
├── auth.json                   # User-supplied
│
└── curriculum_errors/          # Created at runtime
    ├── row_XX_XXXX.png
    └── row_XX_XXXX.json
```

| Technology          | Technical responsibility                              |
| ------------------- | ----------------------------------------------------- |
| Python 3.10+        | Overall execution and orchestration                   |
| Playwright Sync API | Browser automation and DOM interaction                |
| Chromium            | Browser in headless or headed mode                    |
| OpenPyXL            | Read, validate and update Excel                       |
| Dataclasses         | Represent each curriculum node                        |
| Argparse            | Configure execution through terminal flags            |
| Logging             | Record execution events and failures                  |
| JSON                | Read saved authentication state and write diagnostics |
| pathlib             | File paths and directory handling                     |

The packaged dependencies are `playwright>=1.45,<2` and `openpyxl>=3.1,<4`.

The project doesn't use Selenium, direct Extramarks APIs, or a backend database in the original implementation. All creation operations happen through the portal interface.

## 3. Excel input architecture

The Excel file acts as the source of truth for the desired curriculum structure.

The script does not assume a fixed number of levels. It reads a hierarchy path and determines the depth dynamically.

### Excel schema

| Column           | Required | Purpose                                  |
| ---------------- | -------- | ---------------------------------------- |
| `Hierarchy Path` | Yes      | Complete path from Class to the node     |
| `Level Type`     | Yes      | Academic type of node                    |
| `Portal Type`    | No       | Exact option to select in portal         |
| `Description`    | No       | Description entered in the creation form |
| `Create Groups`  | No       | Whether to create three content groups   |
| `Status`         | No       | Execution progress                       |
| `Remarks`        | No       | Error or verification details            |

The expected worksheet is named `Curriculum`. If it is absent, the script uses the active worksheet.

### Example hierarchy

Consider Grade X Science under Philippine Board:

Philippine Board

Existing

X

Class

Science

Subject

Quarter I

Term

Earth and Its Processes

Unit

Plate Tectonics

Chapter

Plate Tectonics: Foundations

Topic

↳ Learning Concepts · Questions Hub · Chapter Tests

The Excel must contain a separate row for each node, including all intermediate parent levels.

For the final Topic, the `Hierarchy Path` would be:

```
X > Science > Quarter I > Earth and Its Processes
> Plate Tectonics > Plate Tectonics: Foundations
```

The script splits this into:

```
(    "X",    "Science",    "Quarter I",    "Earth and Its Processes",    "Plate Tectonics",    "Plate Tectonics: Foundations")
```

It then derives:

```
node.name# "Plate Tectonics: Foundations"node.parent# ("X", "Science", "Quarter I",#  "Earth and Its Processes", "Plate Tectonics")
```

These derived properties are central to the entire automation. The script always navigates using `node.parent`, while the creation function uses `node.name`.

### The `Node` data model

Every Excel row is converted into a dataclass:

```
@dataclassclass Node:    excel_row: int    path: tuple[str, ...]    level_type: str    portal_type: str    description: str    create_groups: bool    status: str    remarks: str
```

This separates raw Excel reading from downstream browser operations.

The original row number is preserved in `excel_row`, even after sorting, so status updates are written to the correct spreadsheet row.

## 4. Excel reading and hierarchy validation

The responsible class is:

```
class ExcelState:
```

It has four main responsibilities.

### A. Workbook loading

```
self.wb = load_workbook(filename)
```

The script loads the complete workbook into memory.

The `val()` method reads values by header name rather than hardcoded column index.

This makes column rearrangement possible without changing Python code, provided the header names remain unchanged.

### B. Normalization

Before comparisons, the code uses:

```
def norm(value):    return unicodedata.normalize(        "NFKC", clean(value)    ).casefold()
```

This performs Unicode normalization, whitespace cleanup and case-insensitive comparison.

For example, `Science`, `science`, and `SCIENCE` are treated as equivalent during hierarchy validation.

This is useful for detecting repeated paths that only differ in capitalization or spacing.

### C. Parent validation

For each hierarchy path, the script verifies that every intermediate parent exists somewhere in Excel.

For example:

`X > Science > Quarter I > Plate Tectonics`

requires rows for `X`, `X > Science`, and `X > Science > Quarter I`.

If any parent is missing, the script stops before opening the portal. It will not invent the parent's level type.

The same validation rejects duplicate normalized hierarchy paths.

### D. Parent-first sorting

The original sorting statement is:

```
return sorted(    nodes,    key=lambda n: (        key_path(n.path),        n.excel_row    ))
```

This groups common hierarchy prefixes and ensures that a parent path precedes its descendants.

It solves the problem of Topic rows appearing before their parents in Excel.

Important limitation: The sort is lexical, not academic sequence-based. Consequently, it can process `Fourth Quarter` before `Second Quarter`. If the Extramarks portal uses creation order as display order, a separate sequence-aware sorter is needed.

## 5. Main execution controller

The central function is:

```
def run(args):
```

This function connects the Excel Manager and Portal automation.

Its execution is divided into six stages.

1. Initialize and validate

   Parse terminal arguments, load the workbook, build the hierarchy plan and identify eligible rows.
2. Initialize Playwright

   Launch Chromium, create a browser context using `auth.json`, and open the Extramarks portal.
3. Navigate to target parent

   Use `Portal.to_parent(node.parent)` to reach the correct hierarchy location.
4. Check or create

   Call `find_child()`. If the card exists, reuse it; otherwise call `create_level()`.
5. Save and verify

   Click Save, search for the created card, and determine the final execution status.
6. Persist progress and continue

   Update the original Excel row, then process the next eligible node.

The simplified core is:

```
for node in pending:    portal.to_parent(node.parent)    existing = portal.find_child(        node.name,        node.level_type    )    if existing is not None:        excel.save_status(            node,            "Already Exists"        )        continue    excel.save_status(node, "In Progress")    save_button = portal.create_level(node)    excel.save_status(        node,        "Failed - Verify"    )    save_button.click()    portal.verify_save(node)    excel.save_status(node, "Created")
```

The actual code also handles dry-run behavior, screenshots, exceptions, stop limits, and browser cleanup.

## 6. Portal navigation engine — Code 1

The `Portal` class is responsible for interacting with the Extramarks website.

### A. Login and Board selection

`enter_syllabus()` follows the portal navigation:

```
Extramarks
    ↓
Account Menu
    ↓
Dashboard
    ↓
Manage
    ↓
Syllabus
    ↓
All Boards → Philippine Board
    ↓
Continue Editing
    ↓
Open Board Card
```

Playwright loads the authenticated browser state through:

```
context = browser.new_context(    storage_state=str(args.auth.resolve()))
```

This avoids writing login credentials into the script, but it depends on a valid existing authenticated session.

### B. Right-panel card detection

One of the most important functions is:

```
def right_cards(self):
```

The original script uses:

```
CARD_SELECTOR = "div.newboard-card-body.radius-8"CARD_NAME_SELECTOR = "p.newboard-card-body-board-name"
```

There are multiple cards in the UI, including a Board catalog on the left and a hierarchy panel on the right.

To distinguish them, the original function calculates the screen midpoint:

```
midpoint = self.page.evaluate(    "() => window.innerWidth / 2")
```

It returns visible cards only when:

```
box["x"] >= midpoint
```

This prevents accidentally selecting a Board card from the left-hand catalog while searching for Classes, Subjects, Units or Topics.

Limitation: This is dependent on the portal layout and viewport width. A genuine hierarchy card may be excluded if its bounding box appears left of the midpoint.

### C. Finding a child card

The key method is:

```
find_child(name, level_type)
```

Its intended algorithm is:

```
Start at top of child-card list
       ↓
Read visible cards
       ↓
Extract name and type
       ↓
Compare against expected node
       ↓
Match found? ── Yes → Return locator
       │
       No
       ↓
Scroll right hierarchy
       ↓
Repeat until found or list exhausted
```

The original function searches using both the card name and its type.

For instance:

```
Expected:
Quarter I / Term

Portal:
Quarter I - Term
```

The base name and type must be reconciled for the match to succeed.

This is where we discovered an important V2 issue: the original suffix parser included `Quarter`, but omitted `Term`.

The corrected regex needs `Term`:

```
TYPE_SUFFIX_RE = re.compile(    r"^(.*?)\s+-\s+"    r"(Board|Class|Subject|Term|Quarter|"    r"Unit|Chapter|Lesson|Topic)\s*$",    re.I)
```

This is a recommended correction; the original archived package did not include it.

### D. Hierarchy scrolling

The `Portal._scroll_right()` function is a separate scrolling mechanism from the group dropdown.

It searches upward through the card's ancestor elements to find a scrollable container, then changes its `scrollTop`.

If direct scrolling does not work, it falls back to:

```
self.page.mouse.wheel(0, 850)
```

with the pointer positioned over a right-panel card.

The purpose is to support lazy-loaded or virtualized hierarchy lists where not every card exists in the rendered DOM simultaneously.

### E. Common-prefix navigation

The navigation state is maintained in:

```
self.current_parent
```

Consider these two consecutive targets:

```
Current:
X > Science > Quarter I > Earth and Its Processes

Next:
X > Science > Quarter I > The Earth's Atmosphere
```

The common prefix is:

```
X > Science > Quarter I
```

Instead of reopening the Board, X, Science and Quarter I, the script returns to the correct child list using breadcrumbs and opens only the different branch.

The primary functions are:

```
to_parent()_to_parent_once()breadcrumb_to_depth()open_child()
```

This reduces navigation operations when processing many Topics under related Chapters.

## 7. Syllabus level creation — Code 2

The responsible function is:

```
create_level(node)
```

It is designed to work with any configured level type, not only Topics.

### Detailed creation sequence

| Stage | Action                   | Playwright implementation          |
| ----- | ------------------------ | ---------------------------------- |
| 1     | Open form                | Click `+ Add New Level`            |
| 2     | Open level-type dropdown | `open_current_type_dropdown()`     |
| 3     | Select required type     | `select_option(node.portal_type)`  |
| 4     | Enter level name         | Fill `Please Enter Syllabus Name*` |
| 5     | Enter description        | Fill `Please Enter Description`    |
| 6     | Configure groups         | Conditional `add_group()` calls    |
| 7     | Locate Save              | Return Save button locator         |
| 8     | Save and verify          | Managed by `run()`                 |

The type is read dynamically from Excel.

For example:

```
Level Type  = Term
Portal Type = term
Name        = Quarter I
```

The script selects `term` and enters `Quarter I`.

Note that level types other than those demonstrated in the legacy source still depend on the live portal supporting the corresponding dropdown options.

### Important implementation detail

`create_level()` does not click Save itself in the original package.

It returns the Save button locator:

```
return save.last
```

The orchestrator controls when Save is clicked. This separation lets the orchestrator update the Excel status to `Failed - Verify` before a potentially non-idempotent Save operation.

That is a deliberate data-safety design.

## 8. Content group creation — legacy Code 2

Group creation is conditional:

```
if node.create_groups:    for name, mode in CONTENT_GROUPS:        self.add_group(name, mode)
```

The configuration is:

| Group Name        | Group Type |
| ----------------- | ---------- |
| Learning Concepts | Learn      |
| Questions Hub     | Practice   |
| Chapter Tests     | Test       |

When `Create Groups` is blank, the code defaults it to Yes for `Topic` and `Lesson`, and No for other level types.

The three groups are added sequentially before the Topic is saved.

### Exact legacy group-selection mechanics

The original legacy script opens Group Name using:

```
placeholder = page.get_by_text(    "Select Group Name",    exact=True).lastdropdown = placeholder.locator(    "xpath=ancestor::div"    "[contains(@class,'emd-select-selector')][1]")dropdown.click()
```

It selects the specific Group Name using direct text locators.

It then opens the Group Type selector with:

```
page.locator(    'input#rc_select_0:visible').last.click()
```

And selects its corresponding value using the legacy text or title locators.

In the original packaged version, the newer `add_group()` calls the generic `select_option()` helper for these selections. During testing, we identified that helper's scroll logic as a potential cause of whole-page scrolling.

The latest recommended group fix restores the legacy direct-selection behavior. I cannot confirm whether that replacement is already present in your remote `v2_curriculum.py`, so it should be treated as a local modification rather than a guaranteed feature of the original ZIP.

## 9. Save verification and duplicate prevention

The system does not assume a successful click means a successful creation.

After saving, it tries to confirm that the new card exists under the expected parent.

The original function is:

```
verify_save(node)
```

It repeatedly calls:

```
found = self.find_child(    node.name,    node.level_type)
```

A successful match allows the script to mark the Excel row as `Created`.

If the card cannot be verified, the operation is considered uncertain.

### Why this matters

Suppose the script creates `Quarter I` but the portal takes several seconds to refresh the right-hand card list.

Without verification, Python might mark it Done too early.

Without the uncertain status, Python might create another Quarter I on the next run.

The package therefore uses the following status logic.

| Status                     | Meaning                      | Next-run behavior                 |
| -------------------------- | ---------------------------- | --------------------------------- |
| Pending / blank            | Not yet processed            | Eligible                          |
| In Progress                | Creation started             | Eligible for retry after checking |
| Already Exists             | Card was found in the portal | Skipped                           |
| Created                    | New card verified            | Skipped                           |
| Failed                     | Error before confirmed Save  | Eligible                          |
| Failed - Verify            | Save may have succeeded      | Skipped unless explicitly retried |
| Done / Completed / Skipped | Considered complete          | Skipped                           |

### Current verification limitation

The original package verifies against the current hierarchy view after Save. It does not automatically reload the full syllabus after each Save.

We subsequently discussed a revised `verify_save()` that re-enters the syllabus and restores the parent path before checking the card. That modification was proposed to handle delayed loading and navigation-state changes, but it is not included in the original archived package.

Also, detecting an existing Topic does not independently confirm that its three content groups are complete.

## 10. Error handling and recovery architecture

The execution is designed to stop on an unexpected error rather than continue blindly.

The package defines custom exceptions:

| Exception               | Purpose                                               |
| ----------------------- | ----------------------------------------------------- |
| `NavigationError`       | General portal navigation or matching failure         |
| `CardNotFound`          | Required existing ancestor could not be located       |
| `DuplicateCardError`    | Multiple matching sibling cards detected              |
| `DryRunMissingAncestor` | Dry run cannot inspect a missing parent's descendants |

### Failure response

When an unexpected node-level exception occurs, the runner:

1. Writes the failure status and error message back to Excel.
2. Captures a full-page screenshot where possible.
3. Writes a JSON diagnostic file.
4. Prints the Python traceback.
5. Stops the run and closes the browser.

The JSON diagnostic file includes information such as:

```
{
  "excel_row": 5,
  "hierarchy": "X > Science > Quarter I",
  "current_parent": ["X", "Science"],
  "url": "Current portal URL",
  "visible_right_cards": [
    "Quarter I - Term"
  ]
}
```

This example illustrates the diagnostic schema; those values are not a claim about a specific saved log.

### What we learned during testing

We encountered three different technical failure categories:

| Problem                            | Likely cause / identified issue                                                        | Corrective direction                                     |
| ---------------------------------- | -------------------------------------------------------------------------------------- | -------------------------------------------------------- |
| `Quarter I` not found              | `- Term` suffix missing from original label parser; possible navigation delay          | Fix label matching and revalidate the parent view        |
| Group dropdown scrolls incorrectly | Generic popup scroll may target page instead of list                                   | Restore legacy direct selection for groups               |
| XLSX corruption                    | Workbook became unreadable; repeated editing through Excel Lite was strongly suspected | Atomic saves, separate progress storage, or safer editor |

These are independent problems. A correct dropdown selector won't fix workbook corruption, and a successful Excel read won't fix portal-card matching.

## 11. Execution modes and command-line controls

The original package supports the following modes.

| Mode                         | Purpose                                      | Portal writes? | Status writes? |
| ---------------------------- | -------------------------------------------- | -------------- | -------------- |
| `--plan-only`                | Validate and print build plan                | No             | No             |
| `--dry-run --headed`         | Inspect portal navigation and existing nodes | No             | No             |
| `--headed --stop-after 5`    | Controlled live batch                        | Yes            | Yes            |
| `--headed`                   | Process all eligible nodes                   | Yes            | Yes            |
| `--headed --retry-uncertain` | Revisit uncertain nodes                      | Potentially    | Yes            |

On your remote Linux environment, using your actual workbook:

```
python3 v2_curriculum.py \
  --excel v2_curriculum_input_RESTORED.xlsx \
  --plan-only
```

For a live small-batch run:

```
python3 v2_curriculum.py \
  --excel v2_curriculum_input_RESTORED.xlsx \
  --headed \
  --stop-after 5
```

The `--stop-after` limit counts processed eligible nodes, including existing nodes. It does not guarantee five new creations.

The `--retry-uncertain` option checks the portal first, but if a card is not detected, the original code can attempt creation. Because we saw card-detection problems, this option should be used only after verifying the current card-matching logic.

Other supported settings include `--sheet`, `--auth`, `--board-card`, `--board-filter`, `--account-button`, `--max-scrolls`, `--verify-timeout`, and `--errors`.

## 12. How one Topic is processed end to end

Take the example:

X → Science → Quarter I → Earth and Its Processes → Plate Tectonics → Plate Tectonics: Foundations

Input stage

ExcelState loads the row and creates a `Node` with the complete path, Topic type, description and group configuration.

Navigation stage

The Portal navigator compares the target parent against `current_parent`. It reuses shared branches and opens only the additional required levels.

Existence check

It searches for `Plate Tectonics: Foundations - Topic` among the expected parent's child cards. If found, creation is skipped.

Creation stage

If missing, it clicks Add New Level, selects `topic`, enters the title and description, then adds Learning Concepts, Questions Hub and Chapter Tests.

Persistence stage

The runner records the uncertain status before clicking Save, then invokes the verification method.

Completion stage

When the saved card is found under the expected parent, the Excel row is marked Created. The runner moves to the next node.

This entire process is repeated for each eligible node until the curriculum is completed or the runner encounters an error.

## 13. Current architecture assessment

Based on the package and the issues we encountered, this is where the system stands.

| Component                           | Implementation          | Observed / known limitation                            |
| ----------------------------------- | ----------------------- | ------------------------------------------------------ |
| Excel reading                       | Implemented             | Susceptible to malformed XLSX files                    |
| Parent validation                   | Implemented             | Requires all parents explicitly                        |
| Parent-first sorting                | Implemented             | Not academic-sequence aware                            |
| Board navigation                    | Implemented             | Depends on existing portal layout                      |
| Common-prefix navigation            | Implemented             | Browser-state desynchronization possible               |
| Hierarchy scrolling                 | Implemented             | Virtualized-list behavior can vary                     |
| Card name and type matching         | Implemented             | Original parser omitted `Term`                         |
| Dynamic level creation              | Implemented             | Portal dropdown types require validation               |
| Group creation                      | Implemented             | Legacy replacement recommended                         |
| Save verification                   | Implemented             | Original version assumes current child list is correct |
| Failure screenshots and diagnostics | Implemented             | No automatic diagnosis                                 |
| Resume by Excel status              | Implemented             | Depends on workbook integrity                          |
| Atomic Excel saving                 | Not in original package | Recommended                                            |
| SQLite progress tracking            | Not implemented         | Proposed for V3                                        |
| PDF content population              | Not implemented         | Potential future integration                           |

### Recommended next architecture

Once the remaining selector issues are stable, I would move to this structure:

```
curriculum_creator_v3/
│
├── main.py
│
├── config.py
│
├── core/
│   ├── excel_reader.py
│   ├── hierarchy_planner.py
│   ├── progress_manager.py
│   └── orchestrator.py
│
├── portal/
│   ├── authentication.py
│   ├── navigator.py
│   ├── level_creator.py
│   ├── group_creator.py
│   └── verifier.py
│
├── input/
│   └── curriculum_input.xlsx
│
├── output/
│   ├── curriculum_progress.sqlite
│   └── curriculum_report.xlsx
│
└── logs/
    ├── execution.log
    ├── screenshots/
    └── diagnostics/
```

The primary advantage would be separation of concerns. Reading the curriculum, navigating the portal, creating levels, managing groups, verifying persistence and tracking progress would no longer depend on a single large script.

In particular, keeping execution progress in SQLite would prevent the automation from repeatedly rewriting its input workbook.

## 14. Final summary

The package is best understood as an Excel-driven, state-aware hierarchy automation engine.

Its four fundamental design ideas are:

| Design principle               | Why it matters                                                                   |
| ------------------------------ | -------------------------------------------------------------------------------- |
| Declarative input              | Excel specifies what hierarchy should exist, not how to click through the portal |
| Parent-aware execution         | Every node is created under the correct existing or newly created parent         |
| Stateful navigation            | The system tracks the current hierarchy and reuses common paths                  |
| Verification before completion | A Save click alone is not considered proof that creation succeeded               |

The most important technical distinction is between the desired state defined by Excel and the observed state discovered through Playwright. The script attempts to reconcile the two.

That design can eventually support many curriculum structures, Boards and subjects without rewriting the automation for every new country.

For now, the main engineering priorities are to make the portal selectors deterministic, handle post-Save navigation reliably, preserve academic sibling order, and protect execution progress from workbook corruption.

Download the original Curriculum Creator package

Source basis: the original packaged Python script, its README/template, the V2 code you shared, and the legacy group-creation script. Changes discussed during debugging are distinguished from features confirmed in the original package.
