using OneNote2Md.ComClient;
using OneNote2Md.Renderer;

namespace OneNote2Md;

/// <summary>
/// Walks a resolved OneNote subtree and writes it to disk as a mirrored folder structure of
/// Markdown files, extracting page images alongside each file.
/// </summary>
public sealed class Exporter
{
    private readonly OneNoteComClient _client;
    private readonly MarkdownRenderer _renderer;
    private readonly ConversionOptions _options;
    private readonly Action<string> _log;
    private readonly bool _showInlineProgress;

    private string _outputDir = string.Empty;
    private int _pageCount;
    private int _imageCount;
    private int _attachmentCount;
    private int _linkCount;
    private int _totalPages;
    private readonly string _indexBaseName;

    // Maps (section .one URL, page name) -> absolute output .md path for every exported page, so
    // intra-export OneNote links can be rewritten to relative Markdown paths.
    private Dictionary<(string OneFile, string PageName), string> _linkIndex = new();

    /// <summary>Initializes a new exporter.</summary>
    /// <param name="client">OneNote COM client used to fetch page content.</param>
    /// <param name="renderer">Renderer that converts page XML to Markdown.</param>
    /// <param name="options">Conversion settings controlling output.</param>
    /// <param name="log">Callback used for verbose and dry-run progress messages.</param>
    public Exporter(OneNoteComClient client, MarkdownRenderer renderer, ConversionOptions options, Action<string> log)
    {
        _client = client;
        _renderer = renderer;
        _options = options;
        _log = log;

        _indexBaseName = Path.GetFileNameWithoutExtension(options.IndexName);
        if (string.IsNullOrWhiteSpace(_indexBaseName)) _indexBaseName = "_index";

        // Per-page progress lines are for non-verbose runs; verbose/dry-run use _log instead.
        _showInlineProgress = !options.Verbose && !options.DryRun;
    }

    // Per-directory naming state shared by every page written into that directory: a running image
    // counter (keeps image file names unique across pages in the folder) and the set of attachment
    // file names already claimed there.
    private sealed class FolderAssets
    {
        public int ImageCounter { get; set; }
        public HashSet<string> AttachmentNames { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Counts the sections (including nested ones) contained in a node.</summary>
    /// <param name="node">The node to count within.</param>
    /// <returns>The total number of sections under (and including) the node.</returns>
    public static int CountSections(OneNoteNode node) => node switch
    {
        SectionNode => 1,
        SectionGroupNode group => group.Sections.Count + group.Groups.Sum(CountSections),
        _ => 0,
    };

    /// <summary>Counts the pages (including sub-pages) contained in a node.</summary>
    /// <param name="node">The node to count within.</param>
    /// <returns>The total number of pages under the node.</returns>
    public static int CountPages(OneNoteNode node) => node switch
    {
        SectionNode section => section.Pages.Sum(CountPagesInPage),
        SectionGroupNode group => group.Sections.Sum(CountPages) + group.Groups.Sum(CountPages),
        PageNode page => CountPagesInPage(page),
        _ => 0,
    };

    private static int CountPagesInPage(PageNode page) => 1 + page.SubPages.Sum(CountPagesInPage);

    /// <summary>Exports the given root node into the configured output directory.</summary>
    /// <param name="root">The resolved section or section group to export.</param>
    /// <param name="sourcePath">Hierarchical display path of the root, used in progress and metadata.</param>
    public void Export(OneNoteNode root, string sourcePath)
    {
        var outputDir = Path.GetFullPath(_options.OutputDirectory);
        _outputDir = outputDir;

        if (!_options.DryRun)
        {
            Directory.CreateDirectory(outputDir);
            if (!_options.Overwrite && Directory.EnumerateFileSystemEntries(outputDir).Any())
            {
                throw new IOException(
                    $"Output directory '{outputDir}' already contains files. Use --overwrite to write into it.");
            }
        }

        _totalPages = CountPages(root);
        _linkIndex = BuildLinkIndex(root, outputDir);

        Console.WriteLine();
        ConsoleEx.WriteLine(ConsoleColor.Yellow, "{0} {1} page(s) to {2}...",
            _options.DryRun ? "Planning" : "Exporting", _totalPages, outputDir);

        switch (root)
        {
            case SectionNode section:
                ExportSectionPages(section, outputDir, sourcePath);
                break;
            case SectionGroupNode group:
                ExportGroup(group, outputDir, sourcePath);
                break;
            case PageNode page:
                ExportPageSubtree(page, outputDir, sourcePath);
                break;
        }

        ConsoleEx.WriteLine("Processed {0} page(s){1}{2}, {3} link(s) rewritten{4}.",
            _pageCount,
            _options.Images == ImageMode.Extract ? $", {_imageCount} image(s)" : string.Empty,
            _attachmentCount > 0 ? $", {_attachmentCount} attachment(s)" : string.Empty,
            _linkCount,
            _options.DryRun ? " (dry run, nothing written)" : string.Empty);
    }

    private void ExportGroup(SectionGroupNode group, string dir, string sourcePath)
    {
        var used = new HashSet<string>();

        foreach (var section in group.Sections)
        {
            var folderName = FileNaming.MakeUnique(FileNaming.ToBaseName(section.Name, _options.FilenameStyle, _options.MaxNameLength), used);
            var folderPath = Path.Combine(dir, folderName);
            EnsureDirectory(folderPath);
            ExportSectionPages(section, folderPath, $"{sourcePath}/{section.Name}");
        }

        foreach (var child in group.Groups)
        {
            var folderName = FileNaming.MakeUnique(FileNaming.ToBaseName(child.Name, _options.FilenameStyle, _options.MaxNameLength), used);
            var folderPath = Path.Combine(dir, folderName);
            EnsureDirectory(folderPath);
            ExportGroup(child, folderPath, $"{sourcePath}/{child.Name}");
        }
    }

    private void ExportSectionPages(SectionNode section, string dir, string sourcePath)
    {
        ExportPages(section.Pages, dir, sourcePath);
    }

    // Exports a single page (and any sub-pages) as the export root, e.g. when resolving a page link.
    private void ExportPageSubtree(PageNode page, string dir, string sourcePath)
    {
        ExportPages(new[] { page }, dir, sourcePath);
    }

    private void ExportPages(IReadOnlyList<PageNode> pages, string dir, string sourcePath)
    {
        var assets = new FolderAssets();

        if (_options.Subpages == SubpageLayout.Flat)
        {
            var used = new HashSet<string>();
            foreach (var page in Flatten(pages))
            {
                var name = FileNaming.MakeUnique(FileNaming.ToBaseName(page.Name, _options.FilenameStyle, _options.MaxNameLength), used);
                WritePage(page, dir, name, PagePrefix(name), $"{sourcePath}/{page.Name}", assets);
            }
        }
        else
        {
            ExportFolderLevel(pages, dir, sourcePath, assets, new HashSet<string>());
        }
    }

    // Writes each page in 'dir'. A page that has sub-pages instead becomes a folder: its own content is
    // written inside that folder as the index file (see --index-name) and its sub-pages are written
    // beside the index (recursively), sharing the folder's image/attachment naming state.
    private void ExportFolderLevel(IReadOnlyList<PageNode> pages, string dir, string sourcePath, FolderAssets assets, HashSet<string> used)
    {
        foreach (var page in pages)
        {
            var baseName = FileNaming.MakeUnique(FileNaming.ToBaseName(page.Name, _options.FilenameStyle, _options.MaxNameLength), used);
            var pageSource = $"{sourcePath}/{page.Name}";

            if (page.SubPages.Count > 0)
            {
                var folderPath = Path.Combine(dir, baseName);
                EnsureDirectory(folderPath);

                var folderAssets = new FolderAssets();
                var folderUsed = new HashSet<string>();
                var indexName = FileNaming.MakeUnique(_indexBaseName, folderUsed);

                WritePage(page, folderPath, indexName, PagePrefix(baseName), pageSource, folderAssets);
                ExportFolderLevel(page.SubPages, folderPath, pageSource, folderAssets, folderUsed);
            }
            else
            {
                WritePage(page, dir, baseName, PagePrefix(baseName), pageSource, assets);
            }
        }
    }

    // First 20 characters of a page's base name, used to make extracted asset file names readable.
    private static string PagePrefix(string baseName)
    {
        var prefix = baseName.Length <= 20 ? baseName : baseName[..20];
        return prefix.TrimEnd(' ', '.');
    }

    private void WritePage(PageNode page, string dir, string fileBaseName, string pagePrefix, string sourcePath, FolderAssets assets)
    {
        _pageCount++;
        var filePath = Path.Combine(dir, fileBaseName + ".md");

        ReportPageProgress(page.Name);

        if (_options.DryRun)
        {
            _log($"  page: {Relative(filePath)}");
        }

        var xml = _client.GetPageXml(page.Id);
        var rendered = _renderer.Render(xml, new RenderContext
        {
            HeadingOffset = _options.HeadingOffset,
            CodeLanguage = _options.CodeLanguage,
            Highlight = _options.Highlight,
            Images = _options.Images,
            FrontMatter = _options.FrontMatter,
            TitleHeading = _options.TitleHeading,
            SourcePath = sourcePath,
            ImageFileAllocator = ext => AllocateImage(assets, pagePrefix, ext),
            AttachmentFileAllocator = name => AllocateAttachment(assets, name),
            LinkResolver = href => ResolveInternalLink(href, filePath),
        });

        if (!_options.DryRun)
        {
            File.WriteAllText(filePath, rendered.Markdown);
            _log($"  wrote: {Relative(filePath)}");
        }

        foreach (var image in rendered.Images)
        {
            _imageCount++;
            var imagePath = Path.Combine(dir, ToLocalPath(image.FileName));
            ReportAssetProgress("Image", image.FileName);
            if (_options.DryRun)
            {
                _log($"    image: {Relative(imagePath)}");
            }
            else
            {
                Directory.CreateDirectory(Path.GetDirectoryName(imagePath)!);
                File.WriteAllBytes(imagePath, image.Data);
                _log($"    wrote image: {Relative(imagePath)}");
            }
        }

        foreach (var attachment in rendered.Attachments)
        {
            _attachmentCount++;
            var attachmentPath = Path.Combine(dir, ToLocalPath(attachment.FileName));
            ReportAssetProgress("Attachment", attachment.FileName);
            if (_options.DryRun)
            {
                _log($"    attachment: {Relative(attachmentPath)}");
            }
            else
            {
                Directory.CreateDirectory(Path.GetDirectoryName(attachmentPath)!);
                File.Copy(attachment.SourceCachePath, attachmentPath, overwrite: true);
                _log($"    wrote attachment: {Relative(attachmentPath)}");
            }
        }
    }

    // Allocates the next unique image file name within a folder: "<NNN>-<page prefix>.<ext>" under the
    // images sub-folder. The per-folder counter guarantees uniqueness regardless of the page prefix.
    private string AllocateImage(FolderAssets assets, string pagePrefix, string ext)
    {
        assets.ImageCounter++;
        var name = $"{assets.ImageCounter:000}-{pagePrefix}.{ext}";
        return $"{_options.ImagesFolder}/{name}";
    }

    // Allocates a unique, sanitized attachment file name (keeping the original preferred name) under
    // the attachments sub-folder.
    private string AllocateAttachment(FolderAssets assets, string preferredName)
    {
        var safe = FileNaming.SanitizeFileName(preferredName);
        var unique = FileNaming.MakeUniqueFileName(safe, assets.AttachmentNames);
        return $"{_options.AttachmentsFolder}/{unique}";
    }

    // Converts a '/'-separated relative asset path into a platform path for combining/writing.
    private static string ToLocalPath(string relativePath) =>
        relativePath.Replace('/', Path.DirectorySeparatorChar);

    // Prints one progress line per page (interactive runs only); verbose/dry-run rely on _log.
    private void ReportPageProgress(string title)
    {
        if (!_showInlineProgress) return;

        ConsoleEx.WriteLine("  Exporting page {0}/{1}: {2}", _pageCount, _totalPages, Truncate(title, 60));
    }

    private static string Truncate(string value, int max) =>
        value.Length <= max ? value : value[..(max - 1)] + "\u2026";

    // Renders a path relative to the output directory so progress output stays compact.
    private string Relative(string path)
    {
        var rel = Path.GetRelativePath(_outputDir, path);
        return rel == "." ? Path.GetFileName(path) : rel;
    }

    private void EnsureDirectory(string path)
    {
        if (_options.DryRun)
        {
            _log($"  dir:  {Relative(path)}");
        }
        else
        {
            Directory.CreateDirectory(path);
        }
    }

    private static IEnumerable<PageNode> Flatten(IEnumerable<PageNode> pages)
    {
        foreach (var page in pages)
        {
            yield return page;
            foreach (var sub in Flatten(page.SubPages))
            {
                yield return sub;
            }
        }
    }

    // ---- intra-export link rewriting ---------------------------------------

    // Resolves a OneNote link href to a relative Markdown path when it targets a page within this
    // export, or returns null to leave the link unchanged. 'currentFile' is the absolute path of the
    // page being rendered, so the result is relative to it (keeping the tree self-contained if moved).
    private string? ResolveInternalLink(string href, string currentFile)
    {
        var target = OneNoteLink.Parse(href);
        if (target?.PageName is null || target.OneFileUrl is null) return null;

        var key = (NormalizeOneFile(target.OneFileUrl), target.PageName.ToLowerInvariant());
        if (!_linkIndex.TryGetValue(key, out var targetFile)) return null;

        var rel = Path.GetRelativePath(Path.GetDirectoryName(currentFile)!, targetFile);
        var mdLink = ToMarkdownLink(rel);

        _linkCount++;
        ReportLinkProgress(target.PageName, rel);

        return mdLink;
    }

    // Reports a rewritten intra-export link beneath the page currently being exported.
    private void ReportLinkProgress(string targetName, string relativePath)
    {
        if (_showInlineProgress)
        {
            ConsoleEx.WriteLine("    Link to: {0}", Truncate(targetName, 60));
        }
        else
        {
            _log($"    link to: {targetName}  ->  {relativePath}");
        }
    }

    // Reports an extracted image or attachment beneath the page currently being exported (inline runs
    // only; verbose/dry-run runs already log asset paths via _log).
    private void ReportAssetProgress(string kind, string fileName)
    {
        if (_showInlineProgress)
        {
            ConsoleEx.WriteLine("    {0}: {1}", kind, Truncate(fileName, 60));
        }
    }

    // Percent-encodes each path segment and joins them with '/' so spaces and characters like
    // parentheses don't break the Markdown link, while '/' stays as the path separator.
    private static string ToMarkdownLink(string relativePath)
    {
        var segments = relativePath.Split('/', '\\');
        return string.Join('/', segments.Select(Uri.EscapeDataString));
    }

    private static string NormalizeOneFile(string url) => url.Replace('\\', '/').ToLowerInvariant();

    // Builds the (section .one URL, page name) -> absolute .md path index by replaying the same
    // naming and folder layout the writer uses, so every link target maps to its real output file.
    private Dictionary<(string, string), string> BuildLinkIndex(OneNoteNode root, string dir)
    {
        var index = new Dictionary<(string, string), string>();
        switch (root)
        {
            case SectionNode section:
                IndexPages(section.Pages, dir, index);
                break;
            case SectionGroupNode group:
                IndexGroup(group, dir, index);
                break;
            case PageNode page:
                IndexPages(new[] { page }, dir, index);
                break;
        }
        return index;
    }

    private void IndexGroup(SectionGroupNode group, string dir, Dictionary<(string, string), string> index)
    {
        var used = new HashSet<string>();

        foreach (var section in group.Sections)
        {
            var folderName = FileNaming.MakeUnique(FileNaming.ToBaseName(section.Name, _options.FilenameStyle, _options.MaxNameLength), used);
            IndexPages(section.Pages, Path.Combine(dir, folderName), index);
        }

        foreach (var child in group.Groups)
        {
            var folderName = FileNaming.MakeUnique(FileNaming.ToBaseName(child.Name, _options.FilenameStyle, _options.MaxNameLength), used);
            IndexGroup(child, Path.Combine(dir, folderName), index);
        }
    }

    private void IndexPages(IReadOnlyList<PageNode> pages, string dir, Dictionary<(string, string), string> index)
    {
        if (_options.Subpages == SubpageLayout.Flat)
        {
            var used = new HashSet<string>();
            foreach (var page in Flatten(pages))
            {
                var name = FileNaming.MakeUnique(FileNaming.ToBaseName(page.Name, _options.FilenameStyle, _options.MaxNameLength), used);
                AddToIndex(index, page, Path.Combine(dir, name + ".md"));
            }
        }
        else
        {
            IndexFolderLevel(pages, dir, index, new HashSet<string>());
        }
    }

    // Mirrors ExportFolderLevel's naming/layout so every page maps to the exact file the writer produces:
    // a page with sub-pages lives at <folder>/<index-name>.md, its sub-pages beside it.
    private void IndexFolderLevel(IReadOnlyList<PageNode> pages, string dir, Dictionary<(string, string), string> index, HashSet<string> used)
    {
        foreach (var page in pages)
        {
            var baseName = FileNaming.MakeUnique(FileNaming.ToBaseName(page.Name, _options.FilenameStyle, _options.MaxNameLength), used);

            if (page.SubPages.Count > 0)
            {
                var folderPath = Path.Combine(dir, baseName);
                var folderUsed = new HashSet<string>();
                var indexName = FileNaming.MakeUnique(_indexBaseName, folderUsed);

                AddToIndex(index, page, Path.Combine(folderPath, indexName + ".md"));
                IndexFolderLevel(page.SubPages, folderPath, index, folderUsed);
            }
            else
            {
                AddToIndex(index, page, Path.Combine(dir, baseName + ".md"));
            }
        }
    }

    private static void AddToIndex(Dictionary<(string, string), string> index, PageNode page, string filePath)
    {
        if (page.OneFilePath is null) return;
        var key = (NormalizeOneFile(page.OneFilePath), page.Name.ToLowerInvariant());
        index.TryAdd(key, filePath);
    }
}
