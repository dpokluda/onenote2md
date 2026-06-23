namespace OneNote2Md.ComClient;

/// <summary>Base type for a node in the exported OneNote subtree.</summary>
public abstract class OneNoteNode
{
    /// <summary>Display name as shown in OneNote (notebook, group, section, or page title).</summary>
    public required string Name { get; init; }
}
