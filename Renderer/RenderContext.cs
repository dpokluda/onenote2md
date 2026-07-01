namespace OneNote2Md.Renderer;

/// <summary>Per-page rendering inputs derived from <see cref="ConversionOptions"/>.</summary>
public sealed class RenderContext
{
    /// <summary>Amount added to OneNote heading levels (OneNote h1 becomes a heading of this depth + 1).</summary>
    public int HeadingOffset { get; init; } = 1;

    /// <summary>Language hint applied to fenced code blocks.</summary>
    public string CodeLanguage { get; init; } = "kql";

    /// <summary>How OneNote text highlighting is rendered in Markdown.</summary>
    public HighlightStyle Highlight { get; init; } = HighlightStyle.Mark;

    /// <summary>Whether embedded images are extracted to disk or dropped.</summary>
    public ImageMode Images { get; init; } = ImageMode.Extract;

    /// <summary>Whether to emit a YAML front-matter block at the top of the page.</summary>
    public bool FrontMatter { get; init; }

    /// <summary>Whether to emit the page title as a top-level heading.</summary>
    public bool TitleHeading { get; init; } = true;

    /// <summary>
    /// Allocates the relative file path (below the page directory, e.g. <c>images/003-Some Page.png</c>)
    /// for the next extracted image with the given extension. The exporter owns naming and de-duplication
    /// so numbering stays unique across every page sharing the folder.
    /// </summary>
    public Func<string, string>? ImageFileAllocator { get; init; }

    /// <summary>
    /// Allocates the relative file path (below the page directory, e.g. <c>attachments/notes.docx</c>)
    /// for an embedded file with the given preferred name, de-duplicating within the folder.
    /// </summary>
    public Func<string, string>? AttachmentFileAllocator { get; init; }

    /// <summary>OneNote source path recorded in front matter.</summary>
    public string SourcePath { get; init; } = string.Empty;

    /// <summary>
    /// Optional callback that rewrites an intra-export OneNote link. Given a raw <c>href</c>, it
    /// returns the relative Markdown link to use, or <c>null</c> to leave the link unchanged.
    /// </summary>
    public Func<string, string?>? LinkResolver { get; init; }
}
