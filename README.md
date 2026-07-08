# onenote2md

A small Windows console application that exports a **OneNote section** (or an entire **section group**) to a tree of local **Markdown files**, preserving the OneNote folder hierarchy and extracting embedded images and file attachments alongside each page.

It talks to the locally installed OneNote desktop client over **COM automation** (the same approach used by our companion `onenote-mcp` project), so no cloud credentials or Graph API access are required; the notebook simply has to be open in OneNote on the same machine.

## How it works

1. Connect to the running OneNote desktop instance via the `Microsoft.Office.Interop.OneNote` COM interop.
2. Resolve the target from either a **path** (`Notebook/Group/Section`) or a OneNote **Copy Link** URL. A link that points at a page exports just that page and its sub-pages; a link that points at a section (or a path) exports the whole section.
3. Walk the subtree:
   - **section groups** → output folders
   - **sections** → output folders
   - **pages** → `.md` files
   - **pages that have sub-pages** → a folder named after the page, whose own content
     is written inside it as an index file (`_index.md` by default) with the sub-pages
     beside it (configurable)
4. For each page, fetch its OneNote XML, render it to Markdown, and write any embedded images into an `images/` sub-folder and any file attachments into an `attachments/` sub-folder next to the page file.
5. Rewrite OneNote links that point at **another page in the same export** into relative Markdown links to that page's `.md` file (links to pages outside the export are left untouched).

The resulting folder tree under the output directory mirrors the OneNote structure beneath the requested path.

## Design choices

- **COM, not Graph.** Reuses the proven COM approach from `onenote-mcp`. All COM-facing helpers live in the `ComClient/` folder and the OneNote single-threaded apartment is serialized behind a lock.
- **Path or link addressing.** Sections are referenced either by their hierarchical display path (`Notebook/Group/.../Section`) or by a OneNote **Copy Link** URL. Links are matched by the backing `.one` file (the section-id/page-id GUIDs in OneNote links do **not** match the COM hierarchy IDs, so the `.one` path and page name are used instead).
- **Hierarchy preserved.** A section group exports as a nested folder tree; a single section exports as one folder of pages. A page that has sub-pages becomes a folder named after the page: the page's own content is written inside as `_index.md` (the leading underscore sorts it to the top) and its sub-pages sit beside it, e.g. `Investment/_index.md` next to `Investment/Short-term safe money assets.md`. The index file name is configurable with `--index-name`.
- **Readable filenames.** Page and folder names default to the OneNote title with spaces preserved (`This is a sample page.md`) because that is the most readable and is handled natively by Obsidian, VS Code, and other Markdown tools. Names are always sanitized for the filesystem (illegal characters removed, reserved names avoided, length capped, duplicates de-duplicated). A `--filename-style` switch offers `kebab` or `snake` for users targeting static-site generators. The length cap is configurable with `--max-name-length` (default 120).
- **Images in an `images/` sub-folder.** Extracted images are written into an `images/` sub-folder next to their page and named `{NNN}-{page name}.{ext}`, where `{NNN}` is a zero-padded counter that is unique *per folder*. The per-folder counter guarantees every image gets a distinct file name (fixing collisions where OneNote reuses image names). The `{page name}` portion is simplified to letters, digits, spaces, `-` and `_` (brackets, parentheses and other punctuation are dropped) so the file name stays clean and the Markdown image link always parses, even for pages whose titles contain characters like `[` or `]`. The sub-folder name is configurable with `--images-folder`. Each extracted image is shown beneath its page in the progress output (`Image: <file name>`).
- **File attachments in an `attachments/` sub-folder.** Files inserted into a page (`<one:InsertedFile>` — zips, `.docx`, `.eml`, `.xml`, etc.) are extracted byte-for-byte into an `attachments/` sub-folder (created only when a page has attachments) and linked inline with their original file name. Extracted files are bit-identical to the originals, so a `.zip` remains a valid archive and a `.docx` remains a valid Word document. The sub-folder name is configurable with `--attachments-folder`. Each extracted attachment is shown beneath its page in the progress output (`Attachment:
  <file name>`).
- **Intra-export links rewritten.** When a page links to another page that is part of the same export, the OneNote link is replaced with a **relative** Markdown link to that page's `.md` file (path segments are percent-encoded, so spaces and characters like parentheses don't break the link). Because the links are relative, the exported tree keeps working if it is moved or copied elsewhere. Targets are matched by the link's backing `.one` file plus the page **name**; links to pages outside the export (or whose stored link text no longer matches the current page name) are left as the original OneNote link. Each rewritten link is shown beneath its page in the progress output (`Link to: <target page>`), and the total is reported in the final summary (`... , N link(s) rewritten`).
- **Nested lists as a flattened outline.** OneNote list nesting (via `<one:OEChildren>`) is rendered as flat paragraph lines with computed, literal outline markers rather than as real Markdown lists. This is deliberate: real Markdown lists require blocks nested inside an item (code fences, images) to be indented to the item's content column, which renderers handle inconsistently and which caused nesting and numbering to break whenever a code block appeared inside a list. With the flattened model each item is an ordinary line, so an interleaved code block or image renders normally and can never restart or swallow the list. Markers are:
  - **Bullets** by depth — `*`, `**`, `***`, … (emitted escaped as `\*` so the asterisks
    show literally instead of creating a Markdown bullet or bold run).
  - **Numbers** as a hierarchical path built from OneNote's per-level markers — `1`,
    `1.1`, `1.a`, `2.3.b`, etc.

  Levels are also lightly indented with non-breaking spaces (`&nbsp;`) so the outline reads as nested without tripping Markdown's indented-code-block rule. Manually typed leading numbers in ordinary paragraphs (pages where the author typed `1.`, `2.` as text rather than using OneNote's list feature) are escaped so Markdown does not renumber them.
- **Inline formatting.** Bold, italic, and **strikethrough** (`~~…~~`, from OneNote's `line-through` runs) are carried over, and compose with links and highlighting.
- **Code blocks with a language.** Code paragraphs are emitted as fenced ```` ``` ```` blocks with a language hint. Detection combines several signals, because OneNote records no semantic "code" marker and authors style snippets inconsistently: a monospace font on the paragraph, KQL pipe lines (and a query's leading table/operator line — even when OneNote glued it to the end of the preceding prose paragraph rather than giving it its own line), and JSON objects (a `{`/`[` opener whose braces balance and that contains a quoted key such as `"name":`, so bracketed prose like `[ARM only]:` is left alone). Adjacent code-like lines are grouped into a single fence. Because the source content is overwhelmingly KQL, the default language is `kql` (overridable); blocks that look like JSON are labelled `json` automatically.
- **Text highlighting.** OneNote stores a highlighter mark as a (non-white) background color on the text run. By default these are emitted as `<mark>…</mark>` (HTML, rendered by GitHub, VS Code, and Obsidian); `--highlight equal` instead emits `==…==` (Obsidian/Pandoc), and `--highlight none` drops the highlight and keeps the plain text.
- **Configurable Markdown.** Front matter, the title heading, heading offset, and sub-page layout are all switchable (see below). The optional YAML front matter includes the page `title`, `created`/`modified` timestamps, `created_by`/`modified_by` authors, the OneNote page id, and the source path. OneNote has no single page-author field, so the authors are derived from paragraph-level authorship — `created_by` from the earliest paragraph and `modified_by` from the most recently edited one — and each is omitted when OneNote did not record it. (The `created`/`modified` timestamps come from the page itself; a page's `modified` time can be newer than its last content edit because OneNote also bumps it on sync.)
- **Command line via `System.CommandLine`.** Arguments and switches are defined and parsed with the `System.CommandLine` NuGet package.

## Requirements

- Windows with the **OneNote desktop** client (2016+) installed.
- The target notebook **open** in OneNote on the same machine.
- [.NET SDK 10.0](https://dotnet.microsoft.com/download) or later (the project targets `net10.0-windows`).

## Build

```sh
dotnet build onenote2md.sln
```

## Usage

```sh
onenote2md --section "<OneNote path>" --output "<folder>" [options] onenote2md --section-link "<OneNote copy-link>" --output "<folder>" [options]
```

### Required arguments

Specify the section to export with **exactly one** of `--section` or `--section-link`, plus `--output`.

| Option                   | Description                                                                 |
| ------------------------ | --------------------------------------------------------------------------- |
| `--section <path>`       | Hierarchical OneNote path to the section or section group to export, e.g. `Governance Vteam Notebook/Policy/On-Call`. |
| `--section-link <url>`   | A OneNote **Copy Link** URL (right-click a section tab or page in OneNote → *Copy Link to ...*). If the link points at a **page**, just that page and its sub-pages are exported; otherwise the whole **section** is exported. The backing `.one` file embedded in the link is matched against the open OneNote hierarchy. Both the `onenote:` protocol link and the SharePoint `Doc.aspx?...` web link are accepted (paste either, or both). |
| `--output <dir>`         | Output folder; its contents mirror the OneNote subtree below the section.   |

### Options

| Option                          | Default     | Description                                                                                  |
| ------------------------------- | ----------- | -------------------------------------------------------------------------------------------- |
| `--images <extract\|skip>`      | `extract`   | Extract embedded images into an `images/` sub-folder next to each page, or drop them. |
| `--filename-style <preserve\|kebab\|snake>` | `preserve` | How page/folder names are derived from OneNote titles.                            |
| `--subpages <folders\|flat>`    | `folders`   | Render a page that has sub-pages as a folder (its own content in the index file, sub-pages beside it), or flatten all pages into one folder. |
| `--images-folder <name>`        | `images`    | Name of the per-directory sub-folder that extracted images are written into.                  |
| `--attachments-folder <name>`   | `attachments` | Name of the per-directory sub-folder that extracted file attachments are written into.       |
| `--index-name <name>`           | `_index.md` | File name used for a parent page's own content inside its sub-page folder.                     |
| `--max-name-length <n>`         | `120`       | Maximum length (characters) of a file/folder name derived from a OneNote title before de-duplication. |
| `--code-language <lang>`        | `kql`       | Language hint applied to fenced code blocks.                                                  |
| `--highlight <mark\|equal\|none>` | `mark`    | How OneNote text highlighting is rendered: `mark` (`<mark>…</mark>`), `equal` (`==…==`), or `none` (plain text). |
| `--front-matter`                | off         | Emit a YAML front-matter block (title, created/modified timestamps and authors, OneNote id, source path). |
| `--no-title-heading`            | off         | Do not emit the page title as a top-level `#` heading (useful with `--front-matter`).         |
| `--heading-offset <n>`          | `1`         | Amount added to OneNote heading levels (OneNote `h1` becomes `#` × (n+1)).                     |
| `--overwrite`                   | off         | Overwrite existing files instead of failing when the output already contains content.         |
| `--dry-run`                     | off         | Print the planned folder/file structure without writing anything.                             |
| `--verbose`                     | off         | Emit detailed progress to the console.                                                         |

### Examples

Export a whole section group, preserving hierarchy and extracting images:

```sh
onenote2md --section "Azure Policy Livesite Handoff/2026 Weekly Summaries" --output ".\export"
```

Export a single TSG section with YAML front matter and no duplicate H1 title:

```sh
onenote2md --section "Governance Vteam Notebook/Policy/On-Call" --output ".\tsg" --front-matter --no-title-heading
```

Export by pasting a OneNote **Copy Link** instead of typing the path. A link to a section exports the whole section; a link to a page exports just that page and its sub-pages:

```sh
# Whole section (link copied from a section tab)
onenote2md --section-link "onenote:https://contoso.sharepoint.com/.../On-Call.one&section-id={...}&end" --output ".\tsg"

# Single page + its sub-pages (link copied from a page)
onenote2md --section-link "onenote:https://contoso.sharepoint.com/.../On-Call.one#TSGs%20for%20runners&page-id={...}&end" --output ".\tsg"
```

Preview the structure without writing files:

```sh
onenote2md --section "Governance Vteam Notebook/Policy" --output ".\out" --dry-run
```

Produce static-site-friendly filenames and skip images:

```sh
onenote2md --section "Governance Vteam Notebook/Policy/On-Call" --output ".\site" --filename-style kebab --images skip
```

## License

See [LICENSE](LICENSE).
