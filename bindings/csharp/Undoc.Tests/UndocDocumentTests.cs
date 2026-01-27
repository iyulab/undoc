using Xunit;

namespace Undoc.Tests;

/// <summary>
/// Basic tests that don't require native library or test files.
/// </summary>
public class BasicTests
{
    [Fact]
    public void MarkdownOptions_HasSensibleDefaults()
    {
        var opts = new MarkdownOptions();
        Assert.False(opts.IncludeFrontmatter);
    }
}

/// <summary>
/// Tests that require the native library.
/// These are skipped in CI where only the managed assembly is available.
/// </summary>
public class NativeLibraryTests
{
    [Fact(Skip = "Requires native library")]
    public void Version_ReturnsNonEmptyString()
    {
        var version = UndocDocument.Version;
        Assert.NotNull(version);
        Assert.NotEmpty(version);
    }

    [Fact(Skip = "Requires native library")]
    public void Version_HasSemverFormat()
    {
        var version = UndocDocument.Version;
        var parts = version.Split('.');
        Assert.True(parts.Length >= 2, "Version should have at least major.minor");
    }

    [Fact(Skip = "Requires native library")]
    public void ParseFile_NonexistentFile_ThrowsFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            UndocDocument.ParseFile("nonexistent.docx"));
    }
}

/// <summary>
/// Integration tests requiring actual Office files.
/// These tests are skipped in CI where test files are not available.
/// </summary>
public class IntegrationTests
{
    private static readonly string TestFilesDir = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory,
        "..", "..", "..", "..", "..", "..", "test-files");

    private static string? GetTestFile(string filename)
    {
        var path = Path.Combine(TestFilesDir, filename);
        return File.Exists(path) ? path : null;
    }

    [Fact(Skip = "Requires native library and test files")]
    public void ParseFile_ValidDocx_ReturnsDocument()
    {
        var path = GetTestFile("file-sample_1MB.docx");
        if (path == null) return;

        using var doc = UndocDocument.ParseFile(path);
        Assert.NotNull(doc);
        Assert.True(doc.SectionCount >= 0);
    }

    [Fact(Skip = "Requires native library and test files")]
    public void ToMarkdown_ReturnsValidMarkdown()
    {
        var path = GetTestFile("file-sample_1MB.docx");
        if (path == null) return;

        using var doc = UndocDocument.ParseFile(path);
        var markdown = doc.ToMarkdown();

        Assert.NotNull(markdown);
        Assert.NotEmpty(markdown);
    }

    [Fact(Skip = "Requires native library and test files")]
    public void ToMarkdown_WithFrontmatter_ContainsFrontmatter()
    {
        var path = GetTestFile("file-sample_1MB.docx");
        if (path == null) return;

        using var doc = UndocDocument.ParseFile(path);
        var options = new MarkdownOptions { IncludeFrontmatter = true };
        var markdown = doc.ToMarkdown(options);

        Assert.Contains("---", markdown);
    }

    [Fact(Skip = "Requires native library and test files")]
    public void ToText_ReturnsValidText()
    {
        var path = GetTestFile("file-sample_1MB.docx");
        if (path == null) return;

        using var doc = UndocDocument.ParseFile(path);
        var text = doc.ToText();

        Assert.NotNull(text);
        Assert.NotEmpty(text);
    }

    [Fact(Skip = "Requires native library and test files")]
    public void ToJson_ReturnsValidJson()
    {
        var path = GetTestFile("file-sample_1MB.docx");
        if (path == null) return;

        using var doc = UndocDocument.ParseFile(path);
        var json = doc.ToJson();

        Assert.NotNull(json);
        Assert.StartsWith("{", json);
    }

    [Fact(Skip = "Requires native library and test files")]
    public void ParseBytes_ValidData_ReturnsDocument()
    {
        var path = GetTestFile("file-sample_1MB.docx");
        if (path == null) return;

        var data = File.ReadAllBytes(path);
        using var doc = UndocDocument.ParseBytes(data);

        Assert.NotNull(doc);
        var markdown = doc.ToMarkdown();
        Assert.NotEmpty(markdown);
    }

    [Fact(Skip = "Requires native library and test files")]
    public void GetResourceIds_ReturnsArray()
    {
        var path = GetTestFile("file-sample_1MB.docx");
        if (path == null) return;

        using var doc = UndocDocument.ParseFile(path);
        var ids = doc.GetResourceIds();

        Assert.NotNull(ids);
    }

    [Fact(Skip = "Requires native library and test files")]
    public void Dispose_CanBeCalledMultipleTimes()
    {
        var path = GetTestFile("file-sample_1MB.docx");
        if (path == null) return;

        var doc = UndocDocument.ParseFile(path);
        doc.Dispose();
        doc.Dispose(); // Should not throw
    }

    [Fact(Skip = "Requires native library and test files")]
    public void AfterDispose_ThrowsObjectDisposedException()
    {
        var path = GetTestFile("file-sample_1MB.docx");
        if (path == null) return;

        var doc = UndocDocument.ParseFile(path);
        doc.Dispose();

        Assert.Throws<ObjectDisposedException>(() => doc.ToMarkdown());
    }
}
