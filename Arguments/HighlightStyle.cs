namespace OneNote2Md;

/// <summary>How OneNote text highlighting (a colored text background) is rendered in Markdown.</summary>
public enum HighlightStyle
{
    /// <summary>Wrap highlighted text in an HTML <c>&lt;mark&gt;</c> element (widely supported).</summary>
    Mark,

    /// <summary>Wrap highlighted text in <c>==</c> markers (Obsidian/Pandoc highlight extension).</summary>
    Equal,

    /// <summary>Drop highlighting and emit the text without any highlight markup.</summary>
    None,
}
