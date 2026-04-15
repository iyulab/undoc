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
}

public class NativeLibraryTests
{
    [Fact]
    public void Version_ReturnsNonEmptyString()
    {
        NativeTestSupport.EnsureNativeLibraryPrepared();

        var version = UndocDocument.Version;

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
}

internal static class NativeTestSupport
{
    private static readonly object Sync = new();
    private static bool _prepared;

    public static void EnsureNativeLibraryPrepared()
    {
        lock (Sync)
        {
            if (_prepared)
                return;

            var builtLibrary = Path.Combine(RepoRoot, "target", "release", NativeLibraryFileName);
            Assert.True(
                File.Exists(builtLibrary),
                $"Expected native library at {builtLibrary}. Run `cargo build --release --features ffi` first.");

            var destination = Path.Combine(AppContext.BaseDirectory, NativeLibraryFileName);
            File.Copy(builtLibrary, destination, overwrite: true);
            _prepared = true;
        }
    }

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

    private static string NativeLibraryFileName =>
        RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "undoc.dll" :
        RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? "libundoc.dylib" :
        "libundoc.so";
}
