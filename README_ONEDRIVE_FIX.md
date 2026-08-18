# Making GIMS safe to run from a laptop-to-laptop OneDrive share

I cannot actually run PyInstaller or produce a real Windows .exe from
this sandbox, since building a Windows executable has to happen on a
real Windows machine with Python and pywin32 installed there (the same
requirement your own README already states for Outlook and native Excel
PivotTables). What I can do, and have done, is fix the actual cause of
"the app corrupts when moved between laptops," and hand you a ready to
use build spec that bakes in your icon.

## What was actually causing corruption

It was not the application code itself. It was four small JSON files
GIMS writes on its own, next to app.py: `signature.json`,
`templates.json`, `run_history.json`, `receivable_run_history.json`, and
`claims_register_run_history.json`.

Two problems, stacked:

1. They were written with a plain `open(path, "w")` followed by
   `json.dump()`. A plain write like that truncates the file to zero
   bytes first, then writes the new content back in. If OneDrive
   happens to sync the file in that exact window (which it can, since
   it watches for any change), it can upload a half-written or entirely
   empty file. Every other laptop then downloads that broken file, and
   the next time GIMS tries to read it, it fails to parse, or silently
   loses history.
2. These files lived inside the same folder as the application code,
   which is exactly the folder being shared across laptops. Two
   laptops both writing their own signature, or both logging a run, at
   close to the same time, means OneDrive is trying to reconcile two
   different versions of the same file. It does not merge them. It
   either picks one and drops the other, or keeps both as a confusingly
   named "conflicted copy," which looks like corruption from the user's
   side either way.

## What I changed

Two new small modules, plus the four files above updated to use them:

- **`app_paths.py`**: moves all four state files out of the shared
  project folder and into a local, per-machine folder
  (`%LOCALAPPDATA%\GIMS` on Windows). Each laptop now keeps its own
  signature and its own run history. There is nothing left for two
  laptops to race over, since they are no longer touching the same
  file at all. The first time this runs on a laptop, it automatically
  copies over whatever old signature/history file already sits next to
  app.py, so nobody's existing signature disappears.
- **`safe_io.py`**: every write now goes to a temporary file first,
  which is flushed and fsynced to disk, then atomically renamed over
  the real file. This is a second, independent layer of protection, on
  top of the first fix, so even a mid-save crash (a forced shutdown,
  Windows Update rebooting the machine) can never leave a half-written
  file behind. Reads are also wrapped so a leftover corrupted file from
  before this fix does not crash the app on launch, it is treated as
  empty history instead.
- **`run_log.py`**, **`receivable_run_log.py`**,
  **`claims_register_run_log.py`**, **`email_generator.py`**: updated
  to use the two modules above. Nothing about their public functions
  (`load_history`, `log_run`, `load_signature`, `save_signature`,
  `load_templates`, `save_template`) changed, so nothing in `app.py` or
  anywhere else that calls them needs to change.

## How to install this fix

1. Copy `app_paths.py` and `safe_io.py` into your project folder,
   alongside `app.py`.
2. Replace your existing `run_log.py`, `receivable_run_log.py`,
   `claims_register_run_log.py`, and `email_generator.py` with the
   versions here.
3. That is it. Nothing else in the project needs to change; every other
   module already only calls the public functions these four expose.

## Building the packaged .exe with your icon

This part has to run on a Windows machine, since PyInstaller builds a
native Windows executable for whatever OS it runs on.

1. Save your icon as a real `.ico` file (not `.png`) at
   `C:\Users\DELL\Downloads\GIMS.ico`. If you only have a `.png` today,
   convert it to `.ico` first, any free converter works, the spec file
   below expects that exact path and filename. If you would rather use
   a different filename or folder, open `GIMS.spec` and edit the
   `icon=` line to match.
2. On that Windows machine, from inside the project folder:
   ```
   pip install -r requirements.txt
   pip install pyinstaller
   pyinstaller GIMS.spec
   ```
3. This produces a `dist\GIMS\` folder containing `GIMS.exe` plus its
   own supporting files. That whole `GIMS` folder, not just the `.exe`
   on its own, is what you put in OneDrive and hand to other laptops.
   Each laptop runs `GIMS.exe` from inside that folder directly.

I used `onedir` mode (a folder of files) rather than `onefile` (a single
self-extracting exe). See the comment block at the top of `GIMS.spec`
for why, in short: a onefile exe re-unpacks itself into a temp folder on
every single launch, which is slow, and both antivirus software and
OneDrive's own scanning tend to flag a large self-extracting exe far
more aggressively than a plain folder of files, which is exactly the
kind of thing that looks like "the app is corrupted" to someone who did
nothing wrong.

## Two more OneDrive habits worth adopting

1. In OneDrive settings, right-click the `GIMS` folder and choose
   "Always keep on this device," rather than leaving it as an
   on-demand, cloud-only placeholder. Otherwise the very first launch
   on a laptop that has not opened the app before can try to run a DLL
   that OneDrive has not actually downloaded yet.
2. Point each module's own "Output folder" setting (Outstanding
   Premium, Receivable, Claims Register, Loss Ratio, UPR Automation,
   Premium Register, Claims Check all have one) at a LOCAL folder
   outside OneDrive, rather than inside the synced project folder,
   especially for a large batch run. Generating hundreds of Excel
   workbooks inside a live-syncing folder means OneDrive is trying to
   upload every single file the instant it is written, which slows the
   run down and adds sync-conflict risk for no benefit, since generated
   reports do not usually need to be shared between laptops the way the
   application code does.

With both the state-file fix and these two habits, the actual mechanism
that was causing "corruption" no longer exists; the application code
never had the problem, the small files sitting next to it did.
