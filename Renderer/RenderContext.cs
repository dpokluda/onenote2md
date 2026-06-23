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

    /// <summary>Sanitized page base name used as the prefix for extracted image files.</summary>
    public required string ImageNamePrefix { get; init; }

    /// <summary>OneNote source path recorded in front matter.</summary>
    public string SourcePath { get; init; } = string.Empty;

    /// <summary>
    /// Optional callback that rewrites an intra-export OneNote link. Given a raw <c>href</c>, it
    /// returns the relative Markdown link to use, or <c>null</c> to leave the link unchanged.
    /// </summary>
    public Func<string, string?>? LinkResolver { get; init; }
}
