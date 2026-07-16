using System.CommandLine;
using System.CommandLine.Invocation;
using OneNote2Md.ComClient;

namespace OneNote2Md;

using Renderer;

/// <summary>
/// Console entry point: defines the command-line interface and runs the export.
/// </summary>
internal static class Program
{
    /// <summary>
    /// Main program entry point.
    /// </summary>
    /// <param name="args">Program arguments.</param>
    /// <returns>The exit code for the invocation.</returns>
    [STAThread]
    private static int Main(string[] args)
    {
        var sectionOption = new Option<string?>("--section",
            "Hierarchical OneNote path to the section or section group to export (e.g. 'Notebook/Group/Section').");
        sectionOption.AddAlias("-s");

        var sectionLinkOption = new Option<string?>("--section-link",
            "A OneNote 'Copy Link' URL to the section to export. Use instead of --section; the backing " +
            ".one file is matched against the OneNote hierarchy.");
        sectionLinkOption.AddAlias("-l");

        var pageOption = new Option<string?>("--page",
            "Hierarchical OneNote path to a single page to export (e.g. 'Notebook/Group/Section/Page Title'). " +
            "The last segment is the page title; any sub-pages are exported too.");
        pageOption.AddAlias("-p");

        var outputOption = new Option<string>("--output",
            "Output folder; its contents mirror the OneNote subtree below --section.")
        { IsRequired = true };
        outputOption.AddAlias("-o");

        var imagesOption = new Option<ImageMode>("--images", () => ImageMode.Extract,
            "Whether to extract embedded images next to each page (extract) or drop them (skip).");

        var filenameStyleOption = new Option<FilenameStyle>("--filename-style", () => FilenameStyle.Preserve,
            "How page/folder names are derived from OneNote titles (preserve, kebab, snake).");

        var subpagesOption = new Option<SubpageLayout>("--subpages", () => SubpageLayout.Folders,
            "Render OneNote sub-pages as nested folders (folders) or flatten them into the section folder (flat).");

        var codeLanguageOption = new Option<string>("--code-language", () => "kql",
            "Language hint applied to fenced code blocks.");

        var highlightOption = new Option<HighlightStyle>("--highlight", () => HighlightStyle.Mark,
            "How OneNote text highlighting is rendered: mark (<mark>...</mark>), equal (==...==), or none.");

        var imagesFolderOption = new Option<string>("--images-folder", () => "images",
            "Sub-folder name (within each page directory) that extracted images are written into.");

        var attachmentsFolderOption = new Option<string>("--attachments-folder", () => "attachments",
            "Sub-folder name (within each page directory) that extracted file attachments are written into.");

        var indexNameOption = new Option<string>("--index-name", () => "_index.md",
            "File name for a page's own content when it has sub-pages (stored inside the page's folder).");

        var maxNameLengthOption = new Option<int>("--max-name-length", () => 120,
            "Maximum length of a generated file/folder base name (before de-duplication suffixes).");

        var frontMatterOption = new Option<bool>("--front-matter",
            "Emit a YAML front-matter block (title, timestamps, authors, OneNote id, source path).");

        var noTitleHeadingOption = new Option<bool>("--no-title-heading",
            "Do not emit the page title as a top-level heading.");

        var headingOffsetOption = new Option<int>("--heading-offset", () => 1,
            "Amount added to OneNote heading levels (OneNote h1 becomes '#' repeated offset+1 times).");

        var overwriteOption = new Option<bool>("--overwrite",
            "Overwrite an output folder that already contains files instead of failing.");

        var dryRunOption = new Option<bool>("--dry-run",
            "Print the planned folder/file structure without writing anything.");

        var verboseOption = new Option<bool>("--verbose",
            "Emit detailed progress to the console.");

        var root = new RootCommand("Export a OneNote section or section group to local Markdown files.")
        {
            sectionOption, sectionLinkOption, pageOption, outputOption, imagesOption, filenameStyleOption, subpagesOption,
            codeLanguageOption, frontMatterOption, noTitleHeadingOption, headingOffsetOption,
            overwriteOption, dryRunOption, verboseOption, highlightOption,
            imagesFolderOption, attachmentsFolderOption, indexNameOption, maxNameLengthOption,
        };

        root.SetHandler((InvocationContext context) =>
        {
            var parsed = context.ParseResult;
            var options = new ConversionOptions
            {
                SectionPath = parsed.GetValueForOption(sectionOption),
                SectionLink = parsed.GetValueForOption(sectionLinkOption),
                PagePath = parsed.GetValueForOption(pageOption),
                OutputDirectory = parsed.GetValueForOption(outputOption)!,
                Images = parsed.GetValueForOption(imagesOption),
                FilenameStyle = parsed.GetValueForOption(filenameStyleOption),
                Subpages = parsed.GetValueForOption(subpagesOption),
                CodeLanguage = parsed.GetValueForOption(codeLanguageOption)!,
                Highlight = parsed.GetValueForOption(highlightOption),
                ImagesFolder = parsed.GetValueForOption(imagesFolderOption)!,
                AttachmentsFolder = parsed.GetValueForOption(attachmentsFolderOption)!,
                IndexName = parsed.GetValueForOption(indexNameOption)!,
                MaxNameLength = parsed.GetValueForOption(maxNameLengthOption),
                FrontMatter = parsed.GetValueForOption(frontMatterOption),
                TitleHeading = !parsed.GetValueForOption(noTitleHeadingOption),
                HeadingOffset = parsed.GetValueForOption(headingOffsetOption),
                Overwrite = parsed.GetValueForOption(overwriteOption),
                DryRun = parsed.GetValueForOption(dryRunOption),
                Verbose = parsed.GetValueForOption(verboseOption),
            };

            context.ExitCode = Run(options);
        });

        return root.Invoke(args);
    }

    private static int Run(ConversionOptions options)
    {
        Console.WriteLine("OneNote2Md: Export a OneNote section or section group to local Markdown files.");
        Action<string> log = options.Verbose || options.DryRun
            ? message => ConsoleEx.WriteLine("{0}", message)
            : _ => { };

        var hasPath = !string.IsNullOrWhiteSpace(options.SectionPath);
        var hasLink = !string.IsNullOrWhiteSpace(options.SectionLink);
        var hasPage = !string.IsNullOrWhiteSpace(options.PagePath);
        if ((hasPath ? 1 : 0) + (hasLink ? 1 : 0) + (hasPage ? 1 : 0) != 1)
        {
            ConsoleEx.WriteLine(ConsoleColor.Red,
                "Error: specify exactly one of --section, --section-link, or --page.");
            return 2;
        }

        try
        {
            Console.WriteLine();
            ConsoleEx.WriteLine(ConsoleColor.Yellow, "Connecting to OneNote...");
            var client = new OneNoteComClient();
            client.Connect();

            OneNoteNode? root;
            string sourcePath;

            if (hasLink)
            {
                Console.WriteLine();
                ConsoleEx.WriteLine(ConsoleColor.Yellow, "Resolving section link...");
                root = client.ResolveTargetByLink(options.SectionLink!, out sourcePath, out var failure);
                if (root is null)
                {
                    ConsoleEx.WriteLine(ConsoleColor.Red, "Error: {0}", failure!);
                    return 2;
                }

                if (root is PageNode page)
                {
                    ConsoleEx.WriteLine("Resolved link to page '{0}/{1}'.", sourcePath, page.Name);
                }
                else
                {
                    ConsoleEx.WriteLine("Resolved link to section '{0}'.", sourcePath);
                }
            }
            else if (hasPage)
            {
                Console.WriteLine();
                ConsoleEx.WriteLine(ConsoleColor.Yellow, "Fetching page '{0}'...", options.PagePath!);
                var page = client.ResolvePageByPath(options.PagePath!, out sourcePath);
                if (page is null)
                {
                    ConsoleEx.WriteLine(ConsoleColor.Red,
                        "Error: Could not find a page at path '{0}'. The path's last segment must be a " +
                        "page title and everything before it a section; make sure the notebook is open " +
                        "in OneNote.", options.PagePath!);
                    return 2;
                }

                root = page;
                ConsoleEx.WriteLine("Resolved page '{0}/{1}'.", sourcePath, page.Name);
            }
            else
            {
                sourcePath = options.SectionPath!;
                Console.WriteLine();
                ConsoleEx.WriteLine(ConsoleColor.Yellow, "Fetching section '{0}'...", sourcePath);
                root = client.ResolveTarget(sourcePath);
                if (root is null)
                {
                    // Be forgiving: the path may actually point at a page. Fall back to a page export
                    // rather than failing, so a page path passed to --section still works.
                    var page = client.ResolvePageByPath(sourcePath, out var pageSectionPath);
                    if (page is not null)
                    {
                        root = page;
                        sourcePath = pageSectionPath;
                        ConsoleEx.WriteLine(
                            "Note: '{0}' is a page, not a section; exporting it as a single page. " +
                            "(Use --page to be explicit.)", options.SectionPath!);
                    }
                    else
                    {
                        ConsoleEx.WriteLine(ConsoleColor.Red,
                            "Error: Could not find a section or section group at path '{0}'. " +
                            "Check the path and make sure the notebook is open in OneNote. " +
                            "(To export a single page, use --page.)", sourcePath);
                        return 2;
                    }
                }
            }

            if (root is PageNode)
            {
                ConsoleEx.WriteLine("Found {0} page(s) (including sub-pages).", Exporter.CountPages(root));
            }
            else
            {
                ConsoleEx.WriteLine("Found {0} section(s) and {1} page(s).",
                    Exporter.CountSections(root), Exporter.CountPages(root));
            }

            var exporter = new Exporter(client, new MarkdownRenderer(), options, log);
            exporter.Export(root, sourcePath);

            Console.WriteLine();
            ConsoleEx.WriteLine(ConsoleColor.Green, "Finished.");
            return 0;
        }
        catch (Exception ex)
        {
            ConsoleEx.WriteLine(ConsoleColor.Red, "Error: {0}", ex.Message);
            if (options.Verbose) Console.Error.WriteLine(ex);
            return 1;
        }
    }
}
