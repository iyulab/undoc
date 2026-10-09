using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace Undoc;

/// <summary>Options for <see cref="UndocDocument.RenderSection"/>.</summary>
/// <remarks>
/// Text is drawn in faces found in <see cref="FontDirectories"/>, then — with
/// <see cref="SystemFonts"/> — in the system's font directories. No face is bundled: on a host
/// without fonts (a minimal container), pass a directory of font files, or the text is counted
/// in <see cref="RenderGaps.TextRuns"/>.
/// </remarks>
public sealed class RenderSectionOptions
{
    /// <summary>Resolution in dots per inch; a slide point is <c>Dpi / 72</c> pixels. Default 150.</summary>
    public float Dpi { get; init; } = 150f;

    /// <summary>Directories searched, with their subdirectories, for font files (TrueType, OpenType, collections).</summary>
    public IReadOnlyList<string> FontDirectories { get; init; } = [];

    /// <summary>Whether the system's font directories are searched too. Default <c>true</c>.</summary>
    public bool SystemFonts { get; init; } = true;
}

/// <summary>
/// What a rendered section could not show, by kind. All zero means everything the slide asks
/// for was painted; otherwise the rest of the slide was still painted.
/// </summary>
public sealed class RenderGaps
{
    /// <summary>Shapes not drawn: a geometry the renderer does not read (custom geometry).</summary>
    [JsonPropertyName("shapes")]
    public uint Shapes { get; init; }

    /// <summary>Pictures not painted: a format other than PNG or JPEG, or no place on the slide.</summary>
    [JsonPropertyName("images")]
    public uint Images { get; init; }

    /// <summary>
    /// Text runs not painted, or painted only in part: no face has their characters, their
    /// script needs shaping, or the text is vertical.
    /// </summary>
    [JsonPropertyName("text_runs")]
    public uint TextRuns { get; init; }

    /// <summary>Charts not drawn.</summary>
    [JsonPropertyName("charts")]
    public uint Charts { get; init; }

    /// <summary>Tables, diagrams (SmartArt) and other graphic frames not drawn.</summary>
    [JsonPropertyName("graphic_frames")]
    public uint GraphicFrames { get; init; }

    /// <summary>
    /// Fills drawn as a stand-in: a pattern in its foreground color, a rectangular or
    /// shape-following gradient as a radial one, a tiled picture stretched.
    /// </summary>
    [JsonPropertyName("approximated_fills")]
    public uint ApproximatedFills { get; init; }

    /// <summary>Whether anything the slide asked for was left unpainted or approximated.</summary>
    [JsonIgnore]
    public bool IsEmpty =>
        Shapes == 0 && Images == 0 && TextRuns == 0 && Charts == 0 && GraphicFrames == 0
        && ApproximatedFills == 0;
}

/// <summary>A rendered section: a PNG, its size in pixels, and what it could not show.</summary>
public sealed class RenderedSection
{
    /// <summary>The slide as a PNG.</summary>
    public required byte[] Png { get; init; }

    /// <summary>Width in pixels.</summary>
    public int Width { get; init; }

    /// <summary>Height in pixels.</summary>
    public int Height { get; init; }

    /// <summary>What the renderer could not paint.</summary>
    public required RenderGaps Gaps { get; init; }

    /// <summary>
    /// Text runs drawn in a face standing in for the one they ask for (it is not installed, or
    /// lacks their characters): readable, but not the slide's own typeface. Not a gap.
    /// </summary>
    public uint SubstitutedTextRuns { get; init; }
}

/// <summary>Wire shape of the native render report.</summary>
internal sealed class RenderInfoPayload
{
    [JsonPropertyName("width")]
    public int Width { get; init; }

    [JsonPropertyName("height")]
    public int Height { get; init; }

    [JsonPropertyName("gaps")]
    public RenderGaps Gaps { get; init; } = new();

    [JsonPropertyName("substituted_text_runs")]
    public uint SubstitutedTextRuns { get; init; }
}

/// <summary>Wire shape of <see cref="RenderSectionOptions"/>.</summary>
internal sealed class RenderOptionsPayload
{
    [JsonPropertyName("dpi")]
    public float Dpi { get; init; }

    [JsonPropertyName("font_dirs")]
    public IReadOnlyList<string> FontDirs { get; init; } = [];

    [JsonPropertyName("system_fonts")]
    public bool SystemFonts { get; init; }
}
