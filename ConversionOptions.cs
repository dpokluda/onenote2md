namespace OneNote2Md;

/// <summary>
/// All user-controllable conversion settings, populated from the command line.
/// </summary>
public sealed class ConversionOptions
{
    /// <summary>
    /// Hierarchical OneNote path to the section or section group to export.
    /// Mutually exclusive with <see cref="SectionLink"/>.
    /// </summary>
    public string? SectionPath { get; init; }

    /// <summary>
    /// A OneNote "Copy Link" URL identifying the section to export. The tool extracts the backing
    /// <c>.one</c> file reference from the link and matches it against the OneNote hierarchy.
    /// Mutually exclusive with <see cref="SectionPath"/>.
    /// </summary>
    public string? SectionLink { get; init; }

    /// <summary>
    /// Hierarchical OneNote path to a single page to export, e.g. "Notebook/Group/Section/Page Title".
    /// The final segment is the page title; everything before it identifies the containing section.
    /// Any sub-pages of the page are exported too. Mutually exclusive with <see cref="SectionPath"/>
    /// and <see cref="SectionLink"/>.
    /// </summary>
    public string? PagePath { get; init; }

    /// <summary>Root output folder; its contents mirror the children of <see cref="SectionPath"/>.</summary>
    public required string OutputDirectory { get; init; }

    /// <summary>Whether to extract images or skip them.</summary>
    public ImageMode Images { get; init; } = ImageMode.Extract;

    /// <summary>Emit a YAML front-matter block (title, timestamps, authors, OneNote id, source path).</summary>
    public bool FrontMatter { get; init; }

    /// <summary>Emit the page title as a top-level heading. Disable when front matter already carries it.</summary>
    public bool TitleHeading { get; init; } = true;

    /// <summary>Amount added to OneNote heading levels (h1 -&gt; this+1).</summary>
    public int HeadingOffset { get; init; } = 1;

    /// <summary>How sub-pages map onto folders.</summary>
    public SubpageLayout Subpages { get; init; } = SubpageLayout.Folders;

    /// <summary>How file/folder names are derived from OneNote titles.</summary>
    public FilenameStyle FilenameStyle { get; init; } = FilenameStyle.Preserve;

    /// <summary>Default fenced-code-block language used for detected code paragraphs.</summary>
    public string CodeLanguage { get; init; } = "kql";

    /// <summary>How OneNote text highlighting is rendered in Markdown.</summary>
    public HighlightStyle Highlight { get; init; } = HighlightStyle.Mark;

    /// <summary>Sub-folder name (per page directory) that extracted images are written into.</summary>
    public string ImagesFolder { get; init; } = "images";

    /// <summary>Sub-folder name (per page directory) that extracted file attachments are written into.</summary>
    public string AttachmentsFolder { get; init; } = "attachments";

    /// <summary>
    /// File name used for a page's own content when that page has sub-pages and is therefore stored
    /// inside its own folder. The leading underscore of the default sorts it to the top of the folder.
    /// </summary>
    public string IndexName { get; init; } = "_index.md";

    /// <summary>Maximum length of a generated file/folder base name before de-duplication suffixes.</summary>
    public int MaxNameLength { get; init; } = 120;

    /// <summary>Overwrite existing files instead of failing when the output already has content.</summary>
    public bool Overwrite { get; init; }

    /// <summary>Print the planned structure without writing anything.</summary>
    public bool DryRun { get; init; }

    /// <summary>Emit verbose progress to the console.</summary>
    public bool Verbose { get; init; }
}
