using Xunit;

namespace Undoc.Tests;

public class UndocDocumentTests
{
    private static readonly string TestFilesDir = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory,
        "..", "..", "..", "..", "..", "..", "test-files");

    [Fact]
    public void Version_ReturnsNonEmptyString()
    {
        var version = UndocDocument.Version;
        Assert.NotNull(version);
        Assert.NotEmpty(version);
    }

    [Fact]
    public void ParseFile_NonexistentFile_ThrowsFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            UndocDocument.ParseFile("nonexistent.docx"));
    }

    [SkippableFact]
    public void ParseFile_ValidDocx_ReturnsDocument()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        Assert.NotNull(doc);
        Assert.True(doc.SectionCount >= 0);
    }

    [SkippableFact]
    public void ToMarkdown_ReturnsValidMarkdown()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var markdown = doc.ToMarkdown();

        Assert.NotNull(markdown);
        Assert.NotEmpty(markdown);
    }

    [SkippableFact]
    public void ToMarkdown_WithFrontmatter_ContainsFrontmatter()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var options = new MarkdownOptions { IncludeFrontmatter = true };
        var markdown = doc.ToMarkdown(options);

        Assert.Contains("---", markdown);
    }

    [SkippableFact]
    public void ToText_ReturnsValidText()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var text = doc.ToText();

        Assert.NotNull(text);
        Assert.NotEmpty(text);
    }

    [SkippableFact]
    public void ToJson_ReturnsValidJson()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var json = doc.ToJson();

        Assert.NotNull(json);
        Assert.StartsWith("{", json);
    }

    [SkippableFact]
    public void ToJson_Compact_ReturnsCompactJson()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var json = doc.ToJson(compact: true);

        Assert.NotNull(json);
        Assert.DoesNotContain("\n  ", json);
    }

    [SkippableFact]
    public void PlainText_ReturnsValidText()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var text = doc.PlainText();

        Assert.NotNull(text);
    }

    [SkippableFact]
    public void SectionCount_ReturnsNonNegative()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        Assert.True(doc.SectionCount >= 0);
    }

    [SkippableFact]
    public void ResourceCount_ReturnsNonNegative()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        Assert.True(doc.ResourceCount >= 0);
    }

    [SkippableFact]
    public void Title_ReturnsNullOrString()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var title = doc.Title;
        // Title may be null or a string
        Assert.True(title == null || title is string);
    }

    [SkippableFact]
    public void Author_ReturnsNullOrString()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var author = doc.Author;
        // Author may be null or a string
        Assert.True(author == null || author is string);
    }

    [SkippableFact]
    public void ParseBytes_ValidData_ReturnsDocument()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        var data = File.ReadAllBytes(path);
        using var doc = UndocDocument.ParseBytes(data);

        Assert.NotNull(doc);
        var markdown = doc.ToMarkdown();
        Assert.NotEmpty(markdown);
    }

    [SkippableFact]
    public void GetResourceIds_ReturnsArray()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var ids = doc.GetResourceIds();

        Assert.NotNull(ids);
    }

    [SkippableFact]
    public void GetResourceInfo_NonexistentId_ReturnsNull()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var info = doc.GetResourceInfo("nonexistent_id");

        Assert.Null(info);
    }

    [SkippableFact]
    public void GetResourceData_NonexistentId_ReturnsNull()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        using var doc = UndocDocument.ParseFile(path);
        var data = doc.GetResourceData("nonexistent_id");

        Assert.Null(data);
    }

    [SkippableFact]
    public void Dispose_CanBeCalledMultipleTimes()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        var doc = UndocDocument.ParseFile(path);
        doc.Dispose();
        doc.Dispose(); // Should not throw
    }

    [SkippableFact]
    public void AfterDispose_ThrowsObjectDisposedException()
    {
        var path = Path.Combine(TestFilesDir, "file-sample_1MB.docx");
        Skip.IfNot(File.Exists(path), "Test file not available");

        var doc = UndocDocument.ParseFile(path);
        doc.Dispose();

        Assert.Throws<ObjectDisposedException>(() => doc.ToMarkdown());
    }
}
