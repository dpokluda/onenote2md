# onenote2md

A small Windows console application that exports a **OneNote section** (or an entire
**section group**) to a tree of local **Markdown files**, preserving the OneNote
folder hierarchy and extracting embedded images alongside each page.

It talks to the locally installed OneNote desktop client over **COM automation**
(the same approach used by our companion `onenote-mcp` project), so no cloud
credentials or Graph API access are required; the notebook simply has to be open
in OneNote on the same machine.

## How it works

1. Connect to the running OneNote desktop instance via the
   `Microsoft.Office.Interop.OneNote` COM interop.
2. Resolve the target from either a **path** (`Notebook/Group/Section`) or a OneNote
   **Copy Link** URL. A link that points at a page exports just that page and its
   sub-pages; a link that points at a section (or a path) exports the whole section.
3. Walk the subtree:
   - **section groups** → output folders
   - **sections** → output folders
   - **pages** → `.md` files
   - **sub-pages** → a same-named folder *beside* the page's file (configurable)
4. For each page, fetch its OneNote XML, render it to Markdown, and write any
   embedded images next to the page file.
5. Rewrite OneNote links that point at **another page in the same export** into
   relative Markdown links to that page's `.md` file (links to pages outside the
   export are left untouched).

The resulting folder tree under the output directory mirrors the OneNote
structure beneath the requested path.

## Design choices

- **COM, not Graph.** Reuses the proven COM approach from `onenote-mcp`. All
  COM-facing helpers live in the `ComClient/` folder and the OneNote
  single-threaded apartment is serialized behind a lock.
- **Path or link addressing.** Sections are referenced either by their hierarchical
  display path (`Notebook/Group/.../Section`) or by a OneNote **Copy Link** URL. Links
  are matched by the backing `.one` file (the section-id/page-id GUIDs in OneNote links
  do **not** match the COM hierarchy IDs, so the `.one` path and page name are used
  instead).
- **Hierarchy preserved.** A section group exports as a nested folder tree; a single
  section exports as one folder of pages. A page that has sub-pages is written as a
  file with a same-named folder *beside* it holding the sub-pages, e.g.
  `Investment.md` next to `Investment/Short-term safe money assets.md`.
- **Readable filenames.** Page and folder names default to the OneNote title with
  spaces preserved (`This is a sample page.md`) because that is the most readable
  and is handled natively by Obsidian, VS Code, and other Markdown tools. Names are
  always sanitized for the filesystem (illegal characters removed, reserved names
  avoided, length capped, duplicates de-duplicated). A `--filename-style` switch
  offers `kebab` or `snake` for users targeting static-site generators.
- **Images next to the page.** Extracted images are written into the **same folder**
  as their page and named `{PageName}_{ImageName}.{ext}`, falling back to
  `{PageName}_Image{NN}.{ext}` when OneNote provides no descriptive name. Each extracted
  image is shown beneath its page in the progress output (`Image: <file name>`).
- **Intra-export links rewritten.** When a page links to another page that is part of
  the same export, the OneNote link is replaced with a **relative** Markdown link to
  that page's `.md` file (path segments are percent-encoded, so spaces and characters
  like parentheses don't break the link). Because the links are relative, the exported
  tree keeps working if it is moved or copied elsewhere. Targets are matched by the
  link's backing `.one` file plus the page **name**; links to pages outside the export
  (or whose stored link text no longer matches the current page name) are left as the
  original OneNote link. Each rewritten link is shown beneath its page in the progress
  output (`Link to: <target page>`), and the total is reported in the final summary
  (`... , N link(s) rewritten`).
- **Code blocks with a language.** Detected code paragraphs are emitted as fenced
  ```` ``` ```` blocks with a language hint. Because the source content is
  overwhelmingly KQL, the default language is `kql` (overridable).
- **Text highlighting.** OneNote stores a highlighter mark as a (non-white) background
  color on the text run. By default these are emitted as `<mark>…</mark>` (HTML, rendered
  by GitHub, VS Code, and Obsidian); `--highlight equal` instead emits `==…==`
  (Obsidian/Pandoc), and `--highlight none` drops the highlight and keeps the plain text.
- **Configurable Markdown.** Front matter, the title heading, heading offset, and
  sub-page layout are all switchable (see below).
- **Command line via `System.CommandLine`.** Arguments and switches are defined and
  parsed with the `System.CommandLine` NuGet package.

## Requirements

- Windows with the **OneNote desktop** client (2016+) installed.
- The target notebook **open** in OneNote on the same machine.
- [.NET SDK 10.0](https://dotnet.microsoft.com/download) or later
  (the project targets `net10.0-windows`).

## Build

```bash
dotnet build onenote2md.sln
```

## Usage

```bash
onenote2md --section "<OneNote path>" --output "<folder>" [options]
onenote2md --section-link "<OneNote copy-link>" --output "<folder>" [options]
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
| `--images <extract\|skip>`      | `extract`   | Extract embedded images next to each page, or drop them.                                      |
| `--filename-style <preserve\|kebab\|snake>` | `preserve` | How page/folder names are derived from OneNote titles.                            |
| `--subpages <folders\|flat>`    | `folders`   | Render OneNote sub-pages into a same-named folder beside the parent page's file, or flatten all pages into one folder. |
| `--code-language <lang>`        | `kql`       | Language hint applied to fenced code blocks.                                                  |
| `--highlight <mark\|equal\|none>` | `mark`    | How OneNote text highlighting is rendered: `mark` (`<mark>…</mark>`), `equal` (`==…==`), or `none` (plain text). |
| `--front-matter`                | off         | Emit a YAML front-matter block (title, created/modified, OneNote id, source path).            |
| `--no-title-heading`            | off         | Do not emit the page title as a top-level `#` heading (useful with `--front-matter`).         |
| `--heading-offset <n>`          | `1`         | Amount added to OneNote heading levels (OneNote `h1` becomes `#` × (n+1)).                     |
| `--overwrite`                   | off         | Overwrite existing files instead of failing when the output already contains content.         |
| `--dry-run`                     | off         | Print the planned folder/file structure without writing anything.                             |
| `--verbose`                     | off         | Emit detailed progress to the console.                                                         |

### Examples

Export a whole section group, preserving hierarchy and extracting images:

```bash
onenote2md --section "Azure Policy Livesite Handoff/2026 Weekly Summaries" --output ".\export"
```

Export a single TSG section with YAML front matter and no duplicate H1 title:

```bash
onenote2md --section "Governance Vteam Notebook/Policy/On-Call" --output ".\tsg" --front-matter --no-title-heading
```

Export by pasting a OneNote **Copy Link** instead of typing the path. A link to a
section exports the whole section; a link to a page exports just that page and its
sub-pages:

```bash
# Whole section (link copied from a section tab)
onenote2md --section-link "onenote:https://contoso.sharepoint.com/.../On-Call.one&section-id={...}&end" --output ".\tsg"

# Single page + its sub-pages (link copied from a page)
onenote2md --section-link "onenote:https://contoso.sharepoint.com/.../On-Call.one#TSGs%20for%20runners&page-id={...}&end" --output ".\tsg"
```

Preview the structure without writing files:

```bash
onenote2md --section "Governance Vteam Notebook/Policy" --output ".\out" --dry-run
```

Produce static-site-friendly filenames and skip images:

```bash
onenote2md --section "Governance Vteam Notebook/Policy/On-Call" --output ".\site" --filename-style kebab --images skip
```

## License

See [LICENSE](LICENSE).
