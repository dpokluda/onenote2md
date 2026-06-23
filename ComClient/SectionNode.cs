namespace OneNote2Md.ComClient;

/// <summary>A section, which contains pages (with sub-pages nested beneath their parent).</summary>
public sealed class SectionNode : OneNoteNode
{
    /// <summary>Top-level pages in this section (sub-pages are nested under their parent page).</summary>
    public List<PageNode> Pages { get; } = new();
}
