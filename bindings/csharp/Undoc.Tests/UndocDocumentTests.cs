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
            () => "ignored",
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
                () => "native null",
                _ => throw new InvalidOperationException("free should not run")));

        Assert.Equal("Failed to get plain text: native null", ex.Message);
    }

    [Fact]
    public void ParseResourceIdsFromNativeJson_PreservesValidEmptyList()
    {
        var ptr = Marshal.StringToCoTaskMemUTF8("[]");
        var freed = false;

        var resourceIds = UndocDocument.ParseResourceIdsFromNativeJson(
            ptr,
            () => "ignored",
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
                () => "native null",
                _ => throw new InvalidOperationException("free should not run")));

        Assert.Equal("Failed to get resource IDs: native null", ex.Message);
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

    public static byte[] CreateMinimalDocxBytes(string text)
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
                """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                </Relationships>
                """);
            WriteEntry(
                zip,
                "word/document.xml",
                $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:r><w:t>{{text}}</w:t></w:r>
                    </w:p>
                  </w:body>
                </w:document>
                """);
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
