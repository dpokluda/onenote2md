namespace OneNote2Md.Renderer;

/// <summary>
/// A single non-image file embedded in a page (a OneNote <c>InsertedFile</c>), ready to be copied
/// next to the page file. The bytes live in a local cache file that OneNote populates on demand.
/// </summary>
/// <param name="FileName">Relative path (below the page directory) to write the attachment as.</param>
/// <param name="SourceCachePath">Absolute path of OneNote's local cache copy of the file's bytes.</param>
public sealed record ExtractedAttachment(string FileName, string SourceCachePath);
