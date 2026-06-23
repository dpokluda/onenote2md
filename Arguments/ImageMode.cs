namespace OneNote2Md;

/// <summary>How embedded images are handled.</summary>
public enum ImageMode
{
    /// <summary>Write image binaries next to the page markdown and link them.</summary>
    Extract,

    /// <summary>Drop images entirely.</summary>
    Skip,
}
