# PagesToDocx

A macOS shell script that recursively converts Apple Pages (`.pages`) documents
to Microsoft Word (`.docx`) while preserving the folder structure.

If you have a directory tree full of Pages documents and need Word copies of
all of them, this script automates the whole job through Pages itself — no
manual "File → Export To → Word" for every file.

## Features

- **Recursive** — finds every `.pages` file in all subfolders.
- **High-fidelity conversion** — uses the installed Apple Pages application to
  export, so layouts, fonts, images, and tables are converted by Pages itself.
- **Preserves folder structure** — optional output directory mirrors the source
  tree, so documents with the same name in different folders stay separate.
- **Safe by default** — existing `.docx` files are skipped unless `--force` is
  used.
- **Non-destructive** — originals are never modified. Documents already open in
  Pages are left alone.
- **Dry-run mode** — preview what would be converted without touching anything.
- **Handles awkward paths** — spaces, quotes, and other special characters in
  file and folder names.
- **Works with both formats** — modern single-file `.pages` documents and older
  `.pages` package bundles.
- **Atomic output** — exports to a temporary file first, then moves it into
  place, so an interrupted conversion does not leave a partial `.docx`.

## Requirements

- macOS with **Apple Pages** installed (App Store or Apple Creator Studio).
- `bash` (preinstalled on macOS).
- Permission for your terminal to automate Pages the first time you run the
  script (macOS will prompt, or you can enable it under
  **System Settings → Privacy & Security → Automation**).

Tested on macOS 15.8.2 with Pages 15.3.1.

> The script does **not** use LibreOffice or third-party converters. Modern
> Pages files use Apple's IWA format, which Pages itself exports most reliably.

## Installation

```bash
git clone https://github.com/ewweberlin/PagesToDocx.git
cd PagesToDocx
chmod +x pages2docx.sh
```

Or simply download `pages2docx.sh`, make it executable, and run it.

## Usage

```bash
./pages2docx.sh [options] <directory>
```

If `<directory>` is omitted, the current directory is used.

### Options

| Option | Description |
| --- | --- |
| `-o, --output DIR` | Write `.docx` files to `DIR`, mirroring the source folder structure. By default each `.docx` is written next to its `.pages` file. |
| `-f, --force` | Overwrite existing `.docx` files. Without this, existing files are skipped. |
| `-n, --dry-run` | Show what would be converted without converting anything. |
| `-q, --quiet` | Only print errors and the final summary. |
| `-h, --help` | Show help. |

### Examples

Convert everything in place (`.docx` next to each `.pages`):

```bash
./pages2docx.sh ~/Documents
```

Convert into a separate folder, keeping the original directory layout:

```bash
./pages2docx.sh -o ~/Documents-Word ~/Documents
```

Preview the work first:

```bash
./pages2docx.sh --dry-run ~/Documents
```

Re-export everything, overwriting old Word files:

```bash
./pages2docx.sh --force ~/Documents
```

## Example output

```text
$ ./pages2docx.sh -o converted "My Documents"

Using /Applications/Pages.app
converting: Reports/2026/Annual Report.pages
converting: Reports/2026/Q1/Budget.pages
converting: Notes/Meeting Notes.pages

Summary: 3 found, 3 converted, 0 skipped, 0 failed
```

```text
My Documents/                     converted/
├── Reports/                      ├── Reports/
│   └── 2026/                     │   └── 2026/
│       ├── Annual Report.pages   │       ├── Annual Report.docx
│       └── Q1/                   │       └── Q1/
│           └── Budget.pages      │           └── Budget.docx
└── Notes/                        └── Notes/
    └── Meeting Notes.pages           └── Meeting Notes.docx
```

## How it works

1. `find` recursively locates every file or package whose name ends in
   `.pages` (hidden files and folders are skipped).
2. For each document, the destination path is calculated — either next to the
   source or under `--output`, mirroring the relative folder path.
3. The script records the IDs of documents already open in Pages.
4. The source document is opened through LaunchServices (`open -a`), the same
   mechanism Finder uses when you double-click a file. This is more reliable
   than AppleScript's `open` command, which can fail or return `missing value`
   while Pages is launching.
5. The script waits until a new document ID appears, then exports that exact
   document as **Microsoft Word** to a hidden temporary `.docx`.
6. The temporary file is moved into place, and only the document that was
   opened by the script is closed. Any documents you already had open are
   untouched.

Export is wrapped in a 600-second AppleScript timeout to accommodate very large
documents.

## Troubleshooting

**`zsh: permission denied: ./pages2docx.sh`**

The executable bit is missing. This usually happens when the script was
downloaded as a single file instead of cloned. Fix it with:

```bash
chmod +x pages2docx.sh
./pages2docx.sh .
```

You can also run the script through `bash` explicitly:

```bash
bash pages2docx.sh .
```

If the error persists, check the file's permissions and extended attributes:

```bash
ls -l@ pages2docx.sh
```

A folder on an external or network volume mounted with the `noexec` option can
also cause this. Copy the script into your home folder and try again.

**"Operation not permitted" or a modal dialog appears in Pages**

Dismiss any Pages dialogs before running the script, and quit Pages if it is in
a stuck state. Then try again:

```bash
osascript -e 'tell application "Pages" to quit'
./pages2docx.sh .
```

**The script reports that the document did not open**

If the source file is already open in Pages, opening it again may only focus the
existing window instead of creating a new document. Close that document in
Pages and retry. The script deliberately does not touch documents that were
already open.

**Automation permission was denied**

Re-enable it under **System Settings → Privacy & Security → Automation** by
allowing your terminal application to control Pages.

**Password-protected documents**

Pages may show a password prompt that the script cannot answer. Unlock or
remove the password first.

**Downloaded documents are blocked by Gatekeeper**

If the `.pages` files came from the internet, macOS may quarantine them. Open
one manually and approve it, or clear the quarantine attribute:

```bash
xattr -dr com.apple.quarantine /path/to/folder
```

**Pages is installed under a different name**

The script detects Pages by bundle identifier (`com.apple.Pages`) rather than
by the app's display name, so renamed installations such as
`Pages Creator Studio.app` are supported.

## Limitations

- macOS only.
- Only `.pages` documents are converted — not Numbers or Keynote files.
- Conversion fidelity is whatever Pages produces; highly complex layouts may
  shift when opened in Word.
- Interactive dialogs (password prompts, missing fonts, repair warnings) cannot
  be handled automatically.

## Exit status

- `0` — all discovered documents were converted or intentionally skipped.
- non-zero — at least one conversion failed, or a setup/usage error occurred.

The final summary always reports how many files were found, converted, skipped,
and failed.

## Contributing

Issues and pull requests are welcome. Please keep the script dependency-free
(other than macOS and Pages) and test changes against a nested folder tree
containing spaces in file names.

## License

Released under the [MIT License](LICENSE).

Copyright © 2026 ewweberlin.
