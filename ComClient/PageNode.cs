namespace OneNote2Md.ComClient;

/// <summary>A single OneNote page, plus any sub-pages nested under it.</summary>
public sealed class PageNode : OneNoteNode
{
    /// <summary>OneNote internal page ID, used to fetch the page content.</summary>
    public required string Id { get; init; }

    /// <summary>Full URL of the backing <c>.one</c> file of the section that owns this page, used to match intra-export links.</summary>
    public string? OneFilePath { get; init; }

    /// <summary>Page indent level reported by OneNote (1 = top level, &gt;1 = sub-page).</summary>
    public int Level { get; init; } = 1;

    /// <summary>UTC timestamp when the page was created.</summary>
    public DateTimeOffset CreatedUtc { get; init; }

    /// <summary>UTC timestamp when the page was last modified.</summary>
    public DateTimeOffset LastModifiedUtc { get; init; }

    /// <summary>Sub-pages (OneNote pages with a deeper indent level) directly under this page.</summary>
    public List<PageNode> SubPages { get; } = new();
}
