namespace OneNote2Md.Renderer;

/// <summary>Result of rendering one page: its Markdown plus any images and attachments to write.</summary>
/// <param name="Markdown">The rendered Markdown text of the page.</param>
/// <param name="Images">Images extracted from the page that must be written alongside it.</param>
/// <param name="Attachments">File attachments extracted from the page that must be copied alongside it.</param>
public sealed record RenderedPage(
    string Markdown,
    IReadOnlyList<ExtractedImage> Images,
    IReadOnlyList<ExtractedAttachment> Attachments);
