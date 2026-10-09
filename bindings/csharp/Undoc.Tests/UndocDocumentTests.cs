using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using Xunit;

namespace Undoc.Tests;

public class BasicTests
{
    [Fact]
    public void MarkdownOptions_HasSensibleDefaults()
    {
        var opts = new MarkdownOptions();
        Assert.False(opts.IncludeFrontmatter);
        Assert.False(opts.Refine);
    }

    [Fact]
    public void MarkdownOptions_RefineIsSettable()
    {
        var opts = new MarkdownOptions { Refine = true };
        Assert.True(opts.Refine);
    }
}

public class Utf8InteropTests
{
    [Fact]
    public void CopyAndFreeNativeUtf8String_CopiesUtf8BeforeFree()
    {
        var ptr = Marshal.StringToCoTaskMemUTF8("Привет из UTF-8");
        var freed = false;

        var value = UndocDocument.CopyAndFreeNativeUtf8String(ptr, p =>
        {
            Assert.Equal(ptr, p);
            Marshal.FreeCoTaskMem(p);
            freed = true;
        });

        Assert.True(freed);
        Assert.Equal("Привет из UTF-8", value);
    }

    [Fact]
    public void PtrToStringUtf8_DecodesUnicodeContent()
    {
        var ptr = Marshal.StringToCoTaskMemUTF8("Здравствуйте");

        try
        {
            Assert.Equal("Здравствуйте", UndocDocument.PtrToStringUtf8(ptr));
        }
        finally
        {
            Marshal.FreeCoTaskMem(ptr);
        }
    }

    [Fact]
    public void CopyAndFreeRequiredNativeUtf8String_PreservesValidEmptyString()
    {
        var ptr = Marshal.StringToCoTaskMemUTF8(string.Empty);
        var freed = false;

        var value = UndocDocument.CopyAndFreeRequiredNativeUtf8String(
            ptr,
            "Failed to get plain text",
            op => new UndocException($"{op}: ignored"),
            p =>
            {
                Assert.Equal(ptr, p);
                Marshal.FreeCoTaskMem(p);
                freed = true;
            });

        Assert.True(freed);
        Assert.Equal(string.Empty, value);
    }

    [Fact]
    public void CopyAndFreeRequiredNativeUtf8String_ThrowsOnNullPointer()
    {
        var ex = Assert.Throws<UndocException>(() =>
            UndocDocument.CopyAndFreeRequiredNativeUtf8String(
                IntPtr.Zero,
                "Failed to get plain text",
                op => new UndocException($"{op}: native null", UndocErrorKind.ZipArchive),
                _ => throw new InvalidOperationException("free should not run")));

        Assert.Equal("Failed to get plain text: native null", ex.Message);
        Assert.Equal(UndocErrorKind.ZipArchive, ex.Kind);
    }

    [Fact]
    public void ParseResourceIdsFromNativeJson_PreservesValidEmptyList()
    {
        var ptr = Marshal.StringToCoTaskMemUTF8("[]");
        var freed = false;

        var resourceIds = UndocDocument.ParseResourceIdsFromNativeJson(
            ptr,
            op => new UndocException($"{op}: ignored"),
            p =>
            {
                Assert.Equal(ptr, p);
                Marshal.FreeCoTaskMem(p);
                freed = true;
            });

        Assert.True(freed);
        Assert.Empty(resourceIds);
    }

    [Fact]
    public void ParseResourceIdsFromNativeJson_ThrowsOnNullPointer()
    {
        var ex = Assert.Throws<UndocException>(() =>
            UndocDocument.ParseResourceIdsFromNativeJson(
                IntPtr.Zero,
                op => new UndocException($"{op}: native null", UndocErrorKind.MissingComponent),
                _ => throw new InvalidOperationException("free should not run")));

        Assert.Equal("Failed to get resource IDs: native null", ex.Message);
        Assert.Equal(UndocErrorKind.MissingComponent, ex.Kind);
    }
}

public class ErrorKindTests
{
    /// <summary>
    /// A message-only exception did not come from the native library, so it carries no
    /// classification — but it must not read as success either.
    /// </summary>
    [Fact]
    public void MessageOnlyException_IsOther_NotNone()
    {
        var ex = new UndocException("wrapper-side failure");

        Assert.Equal(UndocErrorKind.Other, ex.Kind);
        Assert.NotEqual(UndocErrorKind.None, ex.Kind);
    }

    [Fact]
    public void InnerExceptionConstructor_IsOther()
    {
        var ex = new UndocException("wrapped", new InvalidOperationException("inner"));

        Assert.Equal(UndocErrorKind.Other, ex.Kind);
    }

    /// <summary>
    /// Forward compatibility: a newer native library may report a reason this build has
    /// no name for. The number has to survive rather than throw or collapse.
    /// </summary>
    [Fact]
    public void UnknownKindValue_PassesThroughAndKeepsItsNumber()
    {
        var ex = new UndocException("from the future", (UndocErrorKind)9999);

        Assert.Equal(9999, (int)ex.Kind);
        Assert.Equal("9999", ex.Kind.ToString());
    }

    /// <summary>
    /// The C# numbering is only useful if it agrees with the native ABI, so pin it here
    /// too — these values are what cross the boundary.
    /// </summary>
    [Fact]
    public void Discriminants_MatchTheNativeAbi()
    {
        Assert.Equal(0, (int)UndocErrorKind.None);
        Assert.Equal(1, (int)UndocErrorKind.Other);
        Assert.Equal(2, (int)UndocErrorKind.Io);
        Assert.Equal(3, (int)UndocErrorKind.UnknownFormat);
        Assert.Equal(4, (int)UndocErrorKind.UnsupportedFormat);
        Assert.Equal(5, (int)UndocErrorKind.ZipArchive);
        Assert.Equal(6, (int)UndocErrorKind.XmlParse);
        Assert.Equal(7, (int)UndocErrorKind.InvalidData);
        Assert.Equal(8, (int)UndocErrorKind.MissingComponent);
        Assert.Equal(9, (int)UndocErrorKind.Encoding);
        Assert.Equal(10, (int)UndocErrorKind.StyleNotFound);
        Assert.Equal(11, (int)UndocErrorKind.ResourceNotFound);
        Assert.Equal(12, (int)UndocErrorKind.Encrypted);
        Assert.Equal(13, (int)UndocErrorKind.Render);
        Assert.Equal(100, (int)UndocErrorKind.InvalidArgument);
        Assert.Equal(101, (int)UndocErrorKind.Panic);
        Assert.Equal(102, (int)UndocErrorKind.InvalidOutput);
        Assert.Equal(300, (int)UndocErrorKind.SectionOutOfRange);
    }
}

public class NativeErrorKindTests
{
    /// <summary>
    /// The whole point of the feature, end to end: a damaged container must be
    /// recognisable from the exception without reading its message.
    /// </summary>
    [Fact]
    public void ParseBytes_CorruptedArchive_ReportsZipArchive()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();

        var corrupted = new byte[] { 0x50, 0x4B, 0x03, 0x04 }
            .Concat(Encoding.UTF8.GetBytes("truncated garbage with no central directory"))
            .ToArray();

        var ex = Assert.Throws<UndocException>(() => UndocDocument.ParseBytes(corrupted));

        Assert.Equal(UndocErrorKind.ZipArchive, ex.Kind);
        Assert.NotEmpty(ex.Message);
    }

    [Fact]
    public void ParseBytes_NotAnOfficeDocument_ReportsUnknownFormat()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();

        var ex = Assert.Throws<UndocException>(() =>
            UndocDocument.ParseBytes(Encoding.UTF8.GetBytes("not an office document at all")));

        Assert.Equal(UndocErrorKind.UnknownFormat, ex.Kind);
    }

    /// <summary>
    /// Distinct inputs must land on distinct kinds — otherwise the channel exists but
    /// carries no information a caller could act on.
    /// </summary>
    [Fact]
    public void DifferentFailures_ReportDifferentKinds()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();

        var corrupted = new byte[] { 0x50, 0x4B, 0x03, 0x04 }
            .Concat(Encoding.UTF8.GetBytes("garbage"))
            .ToArray();

        var damaged = Assert.Throws<UndocException>(() => UndocDocument.ParseBytes(corrupted));
        var foreign = Assert.Throws<UndocException>(() =>
            UndocDocument.ParseBytes(Encoding.UTF8.GetBytes("plain text file")));

        Assert.NotEqual(damaged.Kind, foreign.Kind);
    }

    /// <summary>
    /// A successful call must leave no classification behind for the next failure check
    /// to pick up.
    /// </summary>
    [Fact]
    public void SuccessfulCall_LeavesNoRecordedKind()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();

        using var doc = UndocDocument.ParseBytes(
            NativeTestSupport.CreateMinimalDocxBytes("hello"));
        _ = doc.ToMarkdown();

        Assert.Equal(0, NativeMethods.undoc_last_error_kind());
    }
}

public class NativeLibraryTests
{
    [Fact]
    public void Version_LoadsFromShippedRuntimePath()
    {
        var stagedLibrary = NativeTestSupport.EnsureNativeLibraryPrepared();

        var version = UndocDocument.Version;

        Assert.Equal(stagedLibrary, NativeTestSupport.StagedLibraryPath);
        Assert.StartsWith(Path.Combine(AppContext.BaseDirectory, "runtimes"), stagedLibrary);
        Assert.False(File.Exists(Path.Combine(AppContext.BaseDirectory, NativeTestSupport.NativeLibraryFileName)));
        Assert.NotNull(version);
        Assert.NotEmpty(version);
    }

    [Fact]
    public void ParseBytes_GeneratedDocx_PreservesUtf8Text()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();

        using var doc = UndocDocument.ParseBytes(
            NativeTestSupport.CreateMinimalDocxBytes("Привет из C#"));

        Assert.Contains("Привет из C#", doc.ToMarkdown());
        Assert.Contains("Привет из C#", doc.ToText());
    }

    /// <summary>
    /// A picture with an SVG original is one picture: the resource info says which file the
    /// document shows and which is its alternate.
    /// </summary>
    [Fact]
    public void GetResourceInfo_MarksAnSvgOriginalAsAlternateOfItsPicture()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();

        using var doc = UndocDocument.ParseBytes(NativeTestSupport.CreateDocxBytes(
            """<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1" descr="Floor plan"/><a:blip r:embed="rIdPng"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rIdSvg"/></a:ext></a:extLst></a:blip></wp:inline></w:drawing></w:r></w:p>""",
            """
            <Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
            <Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.svg"/>
            """,
            ("word/media/image1.png", "png"),
            ("word/media/image2.svg", "<svg/>")));

        using var png = doc.GetResourceInfo("rIdPng");
        Assert.NotNull(png);
        Assert.Equal("primary", png!.RootElement.GetProperty("role").GetString());
        Assert.Equal("Floor plan", png.RootElement.GetProperty("alt_text").GetString());

        using var svg = doc.GetResourceInfo("rIdSvg");
        Assert.NotNull(svg);
        Assert.Equal("alternate", svg!.RootElement.GetProperty("role").GetString());
        Assert.Equal("rIdPng", svg.RootElement.GetProperty("companion_of").GetString());

        Assert.Contains("Floor plan", doc.ToMarkdown());
    }

    [Fact]
    public void CandidatePaths_Include_Windows_Runtime_Native_UndocDll()
    {
        var paths = NativeMethods.BuildCandidatePaths(
            baseDir: "/base",
            assemblyDir: "/assembly",
            runtimeId: "win-x64",
            fileNames: new[] { "undoc_native.dll", "undoc.dll" });

        Assert.Contains(Path.Combine("/base", "runtimes", "win-x64", "native", "undoc.dll"), paths);
        Assert.Contains(Path.Combine("/assembly", "runtimes", "win-x64", "native", "undoc.dll"), paths);
    }
}

public class NativeLibraryResolverTests
{
    [Fact]
    public void CandidatePaths_PreferShippedRuntimeDirectoryOverLooseWindowsCopies()
    {
        using var sandbox = new TemporaryDirectory();
        var baseDir = Path.Combine(sandbox.Path, "base");
        var assemblyDir = Path.Combine(sandbox.Path, "assembly");
        Directory.CreateDirectory(baseDir);
        Directory.CreateDirectory(assemblyDir);

        var shippedRuntimePath = Path.Combine(baseDir, "runtimes", "win-x64", "native", "undoc.dll");
        var assemblyRuntimePath = Path.Combine(assemblyDir, "runtimes", "win-x64", "native", "undoc.dll");
        var looseBasePath = Path.Combine(baseDir, "undoc_native.dll");
        var looseAssemblyPath = Path.Combine(assemblyDir, "undoc.dll");

        CreatePlaceholderFile(shippedRuntimePath);
        CreatePlaceholderFile(assemblyRuntimePath);
        CreatePlaceholderFile(looseBasePath);
        CreatePlaceholderFile(looseAssemblyPath);

        var candidates = NativeMethods.GetCandidatePaths(
            assemblyDir,
            baseDir,
            "win-x64",
            new[] { "undoc_native.dll", "undoc.dll" });

        Assert.Collection(
            candidates,
            candidate => Assert.Equal(shippedRuntimePath, candidate),
            candidate => Assert.Equal(assemblyRuntimePath, candidate),
            candidate => Assert.Equal(looseBasePath, candidate),
            candidate => Assert.Equal(looseAssemblyPath, candidate));
    }

    private static void CreatePlaceholderFile(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "placeholder");
    }
}

public class RenderSectionTests
{
    /// <summary>72 dpi: a slide point is a pixel, and the default slide is 10 × 7.5 inches.</summary>
    private static readonly RenderSectionOptions At72Dpi = new() { Dpi = 72, SystemFonts = false };

    [Fact]
    public void RenderSection_PaintsASlideToPng_WithItsSizeAndGaps()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();
        using var doc = UndocDocument.ParseBytes(NativeTestSupport.CreateMinimalPptxBytes("Slide text"));

        var section = doc.RenderSection(0, At72Dpi);

        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, section.Png[..8]);
        Assert.Equal(720, section.Width);
        Assert.Equal(540, section.Height);
        // System fonts off and none passed: the slide's one run has no face to draw in.
        Assert.Equal(1u, section.Gaps.TextRuns);
        Assert.False(section.Gaps.IsEmpty);
        Assert.Equal(0u, section.Gaps.Shapes);
        Assert.Equal(0u, section.SubstitutedTextRuns);
    }

    [Fact]
    public void RenderSection_DefaultsTo150Dpi()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();
        using var doc = UndocDocument.ParseBytes(NativeTestSupport.CreateMinimalPptxBytes("Slide text"));

        var section = doc.RenderSection(0);

        Assert.Equal(1500, section.Width);
    }

    [Fact]
    public void RenderSection_IndexTheDocumentDoesNotHave_ReportsSectionOutOfRange()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();
        using var doc = UndocDocument.ParseBytes(NativeTestSupport.CreateMinimalPptxBytes("Slide text"));

        var ex = Assert.Throws<UndocException>(() => doc.RenderSection(1, At72Dpi));

        Assert.Equal(UndocErrorKind.SectionOutOfRange, ex.Kind);
    }

    [Fact]
    public void RenderSection_OfADocumentThatIsNotAPresentation_ReportsUnsupportedFormat()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();
        using var doc = UndocDocument.ParseBytes(NativeTestSupport.CreateMinimalDocxBytes("hello"));

        var ex = Assert.Throws<UndocException>(() => doc.RenderSection(0, At72Dpi));

        Assert.Equal(UndocErrorKind.UnsupportedFormat, ex.Kind);
    }

    [Fact]
    public void RenderSection_AResolutionItCannotDraw_ReportsRender()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();
        using var doc = UndocDocument.ParseBytes(NativeTestSupport.CreateMinimalPptxBytes("Slide text"));

        var ex = Assert.Throws<UndocException>(() => doc.RenderSection(0, new RenderSectionOptions { Dpi = 0 }));

        Assert.Equal(UndocErrorKind.Render, ex.Kind);
    }
}

internal static class NativeTestSupport
{
    private static readonly object Sync = new();
    private static bool _prepared;
    private static string? _stagedLibraryPath;

    public static string EnsureNativeLibraryPrepared()
    {
        lock (Sync)
        {
            if (_prepared)
                return _stagedLibraryPath!;

            var runtimeId = NativeMethods.GetRuntimeIdentifierForCurrentPlatform();
            Assert.False(string.IsNullOrEmpty(runtimeId), "Native test runtime identifier should resolve on supported test platforms.");

            var destination = Path.Combine(
                AppContext.BaseDirectory,
                "runtimes",
                runtimeId!,
                "native",
                NativeLibraryFileName);

            DeleteLooseCopies();

            // CI path: the workflow stages the native library directly at
            // the shipping runtime layout before the tests run.
            // Local-dev path: build target/release/<libname> via
            // `cargo build --release --features ffi`, then stage it here.
            if (!File.Exists(destination))
            {
                var builtLibrary = Path.Combine(RepoRoot, "target", "release", NativeLibraryFileName);
                Assert.True(
                    File.Exists(builtLibrary),
                    $"Native library not found at shipping path ({destination}) or local build ({builtLibrary}). "
                    + "In CI, the bindings workflow stages the library at runtimes/<rid>/native/. "
                    + "Locally, run `cargo build --release --features ffi` first.");

                Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                File.Copy(builtLibrary, destination, overwrite: true);
            }

            _stagedLibraryPath = destination;
            _prepared = true;
            return destination;
        }
    }

    public static string StagedLibraryPath => _stagedLibraryPath ?? string.Empty;

    public static byte[] CreateMinimalDocxBytes(string text) =>
        CreateDocxBytes($"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>", relationships: "");

    /// <summary>
    /// A one-slide PPTX whose slide holds <paramref name="text"/> in one text box. It names no
    /// slide size, so the slide is the format's default 10 × 7.5 inches.
    /// </summary>
    public static byte[] CreateMinimalPptxBytes(string text)
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(
                zip,
                "[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """);
            WriteEntry(
                zip,
                "_rels/.rels",
                """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """);
            WriteEntry(
                zip,
                "ppt/presentation.xml",
                """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                </p:presentation>
                """);
            WriteEntry(
                zip,
                "ppt/_rels/presentation.xml.rels",
                """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                </Relationships>
                """);
            WriteEntry(
                zip,
                "ppt/slides/slide1.xml",
                $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                      <p:txBody><a:bodyPr/><a:p><a:r><a:t>{{text}}</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """);
        }

        return stream.ToArray();
    }

    /// <summary>
    /// A DOCX whose body is <paramref name="bodyXml"/>, whose document relationships are
    /// <paramref name="relationships"/>, and which carries <paramref name="parts"/>.
    /// </summary>
    public static byte[] CreateDocxBytes(
        string bodyXml,
        string relationships,
        params (string Path, string Content)[] parts)
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(
                zip,
                "[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);
            WriteEntry(
                zip,
                "_rels/.rels",
                """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);
            WriteEntry(
                zip,
                "word/_rels/document.xml.rels",
                $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                {{relationships}}
                </Relationships>
                """);
            WriteEntry(
                zip,
                "word/document.xml",
                $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>{{bodyXml}}</w:body>
                </w:document>
                """);
            foreach (var (path, content) in parts)
            {
                WriteEntry(zip, path, content);
            }
        }

        return stream.ToArray();
    }

    private static void WriteEntry(ZipArchive zip, string path, string content)
    {
        var entry = zip.CreateEntry(path, CompressionLevel.NoCompression);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }

    private static string RepoRoot =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", ".."));

    public static string NativeLibraryFileName =>
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "undoc.dll" :
        RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? "libundoc.dylib" :
        "libundoc.so";

    public static string RuntimeIdentifier =>
        NativeMethods.GetRuntimeIdentifier() ??
        throw new PlatformNotSupportedException("No shipped native runtime asset is configured for this platform.");

    private static string NativeRuntimeDirectory =>
        Path.Combine(AppContext.BaseDirectory, "runtimes", RuntimeIdentifier, "native");

    private static string NativeLibraryDestination =>
        Path.Combine(NativeRuntimeDirectory, NativeLibraryFileName);

    private static void DeleteLooseCopies()
    {
        foreach (var fileName in RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                     ? new[] { "undoc_native.dll", "undoc.dll" }
                     : new[] { NativeLibraryFileName })
        {
            var loosePath = Path.Combine(AppContext.BaseDirectory, fileName);
            if (File.Exists(loosePath))
                File.Delete(loosePath);
        }
    }
}

internal sealed class TemporaryDirectory : IDisposable
{
    public TemporaryDirectory()
    {
        Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"undoc-csharp-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(Path);
    }

    public string Path { get; }

    public void Dispose()
    {
        if (Directory.Exists(Path))
            Directory.Delete(Path, recursive: true);
    }
}
