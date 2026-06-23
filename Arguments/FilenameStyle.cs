namespace OneNote2Md;

/// <summary>How OneNote titles are turned into file and folder names.</summary>
public enum FilenameStyle
{
    /// <summary>Keep the title as-is with spaces preserved ("This is a sample page").</summary>
    Preserve,

    /// <summary>Lower-case words joined by hyphens ("this-is-a-sample-page").</summary>
    Kebab,

    /// <summary>Lower-case words joined by underscores ("this_is_a_sample_page").</summary>
    Snake,
}
