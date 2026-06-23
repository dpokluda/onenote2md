namespace OneNote2Md.Renderer;

/// <summary>A single image pulled out of a page, ready to be written next to the page file.</summary>
/// <param name="FileName">File name (including extension) to write the image as.</param>
/// <param name="Data">Raw image bytes.</param>
public sealed record ExtractedImage(string FileName, byte[] Data);
