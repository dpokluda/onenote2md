using System.Text;

namespace OneNote2Md;

/// <summary>
/// Turns OneNote titles into filesystem-safe file and folder names according to the chosen
/// <see cref="FilenameStyle"/>, and de-duplicates names that collide within a folder.
/// </summary>
public static class FileNaming
{
    private const int DefaultMaxLength = 120;

    private static readonly HashSet<string> ReservedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    };

    private static readonly char[] InvalidChars = Path.GetInvalidFileNameChars();

    /// <summary>Produces a sanitized base name (no extension) from a OneNote title.</summary>
    /// <param name="title">The OneNote page, section, or group title to convert.</param>
    /// <param name="style">The naming style to apply.</param>
    /// <param name="maxLength">Maximum length of the returned base name.</param>
    /// <returns>A filesystem-safe base name without an extension.</returns>
    public static string ToBaseName(string title, FilenameStyle style, int maxLength = DefaultMaxLength)
    {
        var name = style switch
        {
            FilenameStyle.Kebab => JoinWords(title, '-'),
            FilenameStyle.Snake => JoinWords(title, '_'),
            _ => Preserve(title),
        };

        if (string.IsNullOrWhiteSpace(name)) name = "untitled";

        if (maxLength > 0 && name.Length > maxLength) name = name[..maxLength].TrimEnd();

        // Windows rejects names ending in a dot or space.
        name = name.TrimEnd('.', ' ');
        if (string.IsNullOrWhiteSpace(name)) name = "untitled";

        if (ReservedNames.Contains(name)) name += "_";

        return name;
    }

    /// <summary>
    /// Returns a unique base name within a folder by appending " (2)", " (3)", ... when the
    /// requested name is already taken.
    /// </summary>
    /// <param name="baseName">The desired base name.</param>
    /// <param name="used">Set tracking names already claimed in the folder (case-insensitive); the chosen name is added to it.</param>
    /// <returns>A name not yet present in <paramref name="used"/>.</returns>
    public static string MakeUnique(string baseName, HashSet<string> used)
    {
        var candidate = baseName;
        var counter = 2;
        while (!used.Add(candidate.ToLowerInvariant()))
        {
            candidate = $"{baseName} ({counter})";
            counter++;
        }
        return candidate;
    }

    /// <summary>Replaces characters invalid in a file name with underscores, preserving the extension.</summary>
    /// <param name="fileName">The desired file name (including extension).</param>
    /// <returns>A filesystem-safe file name.</returns>
    public static string SanitizeFileName(string fileName)
    {
        var sb = new StringBuilder(fileName.Length);
        foreach (var ch in fileName)
        {
            sb.Append(Array.IndexOf(InvalidChars, ch) >= 0 ? '_' : ch);
        }
        var cleaned = CollapseWhitespace(sb.ToString()).Trim().TrimEnd('.', ' ');
        return string.IsNullOrWhiteSpace(cleaned) ? "attachment" : cleaned;
    }

    /// <summary>
    /// Returns a unique file name within a folder by inserting " (2)", " (3)", ... before the
    /// extension when the requested name is already taken.
    /// </summary>
    /// <param name="fileName">The desired file name (including extension).</param>
    /// <param name="used">Set tracking file names already claimed in the folder (case-insensitive).</param>
    /// <returns>A file name not yet present in <paramref name="used"/>.</returns>
    public static string MakeUniqueFileName(string fileName, HashSet<string> used)
    {
        var ext = Path.GetExtension(fileName);
        var stem = Path.GetFileNameWithoutExtension(fileName);
        var candidate = fileName;
        var counter = 2;
        while (!used.Add(candidate.ToLowerInvariant()))
        {
            candidate = $"{stem} ({counter}){ext}";
            counter++;
        }
        return candidate;
    }

    /// <summary>
    /// Reduces a base name to a simple, link-safe token for use in extracted asset (image/attachment)
    /// file names: keeps letters, digits, spaces, <c>-</c> and <c>_</c>, drops everything else (brackets,
    /// parentheses, dots, <c>#</c>, ...), and collapses whitespace. This keeps asset file names clean and
    /// safe to embed inside Markdown <c>](path)</c> links, where characters like <c>[]()</c> break parsing.
    /// </summary>
    /// <param name="baseName">The (already filesystem-safe) base name to simplify.</param>
    /// <returns>A simplified token containing only letters, digits, spaces, <c>-</c> and <c>_</c>.</returns>
    public static string ToAssetPrefix(string baseName)
    {
        var sb = new StringBuilder(baseName.Length);
        foreach (var ch in baseName)
        {
            if (char.IsLetterOrDigit(ch) || ch is ' ' or '-' or '_') sb.Append(ch);
        }
        return CollapseWhitespace(sb.ToString()).Trim();
    }

    private static string Preserve(string title)
    {
        var sb = new StringBuilder(title.Length);
        foreach (var ch in title)
        {
            if (ch is '/' or '\\')
            {
                sb.Append('-');
            }
            else
            {
                sb.Append(Array.IndexOf(InvalidChars, ch) >= 0 ? ' ' : ch);
            }
        }
        return CollapseWhitespace(sb.ToString()).Trim();
    }

    private static string JoinWords(string title, char separator)
    {
        var words = new List<string>();
        var current = new StringBuilder();
        foreach (var ch in title)
        {
            if (char.IsLetterOrDigit(ch))
            {
                current.Append(char.ToLowerInvariant(ch));
            }
            else if (current.Length > 0)
            {
                words.Add(current.ToString());
                current.Clear();
            }
        }
        if (current.Length > 0) words.Add(current.ToString());
        return string.Join(separator, words);
    }

    private static string CollapseWhitespace(string value)
    {
        var sb = new StringBuilder(value.Length);
        var previousWasSpace = false;
        foreach (var ch in value)
        {
            var isSpace = char.IsWhiteSpace(ch);
            if (isSpace)
            {
                if (!previousWasSpace) sb.Append(' ');
            }
            else
            {
                sb.Append(ch);
            }
            previousWasSpace = isSpace;
        }
        return sb.ToString();
    }
}
