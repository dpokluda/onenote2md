namespace OneNote2Md.ComClient;

/// <summary>A section group (or notebook root), which contains sections and nested groups.</summary>
public sealed class SectionGroupNode : OneNoteNode
{
    /// <summary>Nested section groups directly under this group.</summary>
    public List<SectionGroupNode> Groups { get; } = new();

    /// <summary>Sections directly under this group.</summary>
    public List<SectionNode> Sections { get; } = new();
}
