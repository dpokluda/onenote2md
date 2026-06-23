namespace OneNote2Md;

/// <summary>How OneNote sub-pages are mapped onto the output folder structure.</summary>
public enum SubpageLayout
{
    /// <summary>A page with sub-pages becomes a folder containing its own file plus child files.</summary>
    Folders,

    /// <summary>All pages in a section land directly in the section folder, ignoring sub-page nesting.</summary>
    Flat,
}
