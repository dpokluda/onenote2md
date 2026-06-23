namespace OneNote2Md.Renderer;

/// <summary>Result of rendering one page: its Markdown plus any images that must be written.</summary>
/// <param name="Markdown">The rendered Markdown text of the page.</param>
/// <param name="Images">Images extracted from the page that must be written alongside it.</param>
public sealed record RenderedPage(string Markdown, IReadOnlyList<ExtractedImage> Images);
