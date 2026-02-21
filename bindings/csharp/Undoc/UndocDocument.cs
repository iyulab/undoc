using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace Undoc;

/// <summary>
/// Exception thrown when an undoc operation fails.
/// </summary>
public class UndocException : Exception
{
    public UndocException(string message) : base(message) { }
}

/// <summary>
/// Options for markdown rendering.
/// </summary>
public class MarkdownOptions
{
    /// <summary>
    /// Include YAML frontmatter with document metadata.
    /// </summary>
    public bool IncludeFrontmatter { get; set; } = false;

    /// <summary>
    /// Escape special markdown characters.
    /// </summary>
    public bool EscapeSpecialChars { get; set; } = false;

    /// <summary>
    /// Add extra spacing between paragraphs.
    /// </summary>
    public bool ParagraphSpacing { get; set; } = false;

    internal uint ToFlags()
    {
        uint flags = 0;
        if (IncludeFrontmatter) flags |= NativeMethods.UNDOC_FLAG_FRONTMATTER;
        if (EscapeSpecialChars) flags |= NativeMethods.UNDOC_FLAG_ESCAPE_SPECIAL;
        if (ParagraphSpacing) flags |= NativeMethods.UNDOC_FLAG_PARAGRAPH_SPACING;
        return flags;
    }
}

/// <summary>
/// Represents a parsed Office document.
/// </summary>
/// <remarks>
/// This class provides methods to extract content from DOCX, XLSX, and PPTX
/// documents in various formats (Markdown, plain text, JSON).
/// </remarks>
public class UndocDocument : IDisposable
{
    private IntPtr _handle;
    private bool _disposed;

    private UndocDocument(IntPtr handle)
    {
        _handle = handle;
    }

    /// <summary>
    /// Get the undoc library version.
    /// </summary>
    public static string Version
    {
        get
        {
            var ptr = NativeMethods.undoc_version();
            return Marshal.PtrToStringAnsi(ptr) ?? "unknown";
        }
    }

    /// <summary>
    /// Parse a document from a file path.
    /// </summary>
    /// <param name="path">Path to the document file</param>
    /// <returns>Parsed document</returns>
    /// <exception cref="UndocException">If parsing fails</exception>
    /// <exception cref="FileNotFoundException">If file doesn't exist</exception>
    public static UndocDocument ParseFile(string path)
    {
        if (!System.IO.File.Exists(path))
            throw new System.IO.FileNotFoundException($"File not found: {path}", path);

        var handle = NativeMethods.undoc_parse_file(path);
        if (handle == IntPtr.Zero)
            throw new UndocException($"Failed to parse {path}: {GetLastError()}");

        return new UndocDocument(handle);
    }

    /// <summary>
    /// Parse a document from a byte array.
    /// </summary>
    /// <param name="data">Document content as bytes</param>
    /// <returns>Parsed document</returns>
    /// <exception cref="UndocException">If parsing fails</exception>
    public static UndocDocument ParseBytes(byte[] data)
    {
        var dataPtr = Marshal.AllocHGlobal(data.Length);
        try
        {
            Marshal.Copy(data, 0, dataPtr, data.Length);
            var handle = NativeMethods.undoc_parse_bytes(dataPtr, (UIntPtr)data.Length);
            if (handle == IntPtr.Zero)
                throw new UndocException($"Failed to parse bytes: {GetLastError()}");

            return new UndocDocument(handle);
        }
        finally
        {
            Marshal.FreeHGlobal(dataPtr);
        }
    }

    /// <summary>
    /// Convert the document to Markdown.
    /// </summary>
    /// <param name="options">Optional rendering options</param>
    /// <returns>Markdown string</returns>
    public string ToMarkdown(MarkdownOptions? options = null)
    {
        ThrowIfDisposed();
        uint flags = options?.ToFlags() ?? 0;
        var ptr = NativeMethods.undoc_to_markdown(_handle, flags);
        if (ptr == IntPtr.Zero)
            throw new UndocException($"Failed to convert to markdown: {GetLastError()}");

        try
        {
            return PtrToStringUtf8(ptr);
        }
        finally
        {
            NativeMethods.undoc_free_string(ptr);
        }
    }

    /// <summary>
    /// Convert the document to plain text.
    /// </summary>
    /// <returns>Plain text string</returns>
    public string ToText()
    {
        ThrowIfDisposed();
        var ptr = NativeMethods.undoc_to_text(_handle);
        if (ptr == IntPtr.Zero)
            throw new UndocException($"Failed to convert to text: {GetLastError()}");

        try
        {
            return PtrToStringUtf8(ptr);
        }
        finally
        {
            NativeMethods.undoc_free_string(ptr);
        }
    }

    /// <summary>
    /// Convert the document to JSON.
    /// </summary>
    /// <param name="compact">Use compact JSON format</param>
    /// <returns>JSON string</returns>
    public string ToJson(bool compact = false)
    {
        ThrowIfDisposed();
        int format = compact ? NativeMethods.UNDOC_JSON_COMPACT : NativeMethods.UNDOC_JSON_PRETTY;
        var ptr = NativeMethods.undoc_to_json(_handle, format);
        if (ptr == IntPtr.Zero)
            throw new UndocException($"Failed to convert to JSON: {GetLastError()}");

        try
        {
            return PtrToStringUtf8(ptr);
        }
        finally
        {
            NativeMethods.undoc_free_string(ptr);
        }
    }

    /// <summary>
    /// Get plain text content (faster than ToText for simple extraction).
    /// </summary>
    /// <returns>Plain text string</returns>
    public string PlainText()
    {
        ThrowIfDisposed();
        var ptr = NativeMethods.undoc_plain_text(_handle);
        if (ptr == IntPtr.Zero)
            throw new UndocException($"Failed to get plain text: {GetLastError()}");

        try
        {
            return PtrToStringUtf8(ptr);
        }
        finally
        {
            NativeMethods.undoc_free_string(ptr);
        }
    }

    /// <summary>
    /// Get the number of sections in the document.
    /// </summary>
    public int SectionCount
    {
        get
        {
            ThrowIfDisposed();
            var count = NativeMethods.undoc_section_count(_handle);
            if (count < 0)
                throw new UndocException($"Failed to get section count: {GetLastError()}");
            return count;
        }
    }

    /// <summary>
    /// Get the number of resources in the document.
    /// </summary>
    public int ResourceCount
    {
        get
        {
            ThrowIfDisposed();
            var count = NativeMethods.undoc_resource_count(_handle);
            if (count < 0)
                throw new UndocException($"Failed to get resource count: {GetLastError()}");
            return count;
        }
    }

    /// <summary>
    /// Get the document title, if set.
    /// </summary>
    public string? Title
    {
        get
        {
            ThrowIfDisposed();
            var ptr = NativeMethods.undoc_get_title(_handle);
            if (ptr == IntPtr.Zero)
                return null;

            try
            {
                return PtrToStringUtf8(ptr);
            }
            finally
            {
                NativeMethods.undoc_free_string(ptr);
            }
        }
    }

    /// <summary>
    /// Get the document author, if set.
    /// </summary>
    public string? Author
    {
        get
        {
            ThrowIfDisposed();
            var ptr = NativeMethods.undoc_get_author(_handle);
            if (ptr == IntPtr.Zero)
                return null;

            try
            {
                return PtrToStringUtf8(ptr);
            }
            finally
            {
                NativeMethods.undoc_free_string(ptr);
            }
        }
    }

    /// <summary>
    /// Get list of resource IDs in the document.
    /// </summary>
    /// <returns>Array of resource ID strings</returns>
    public string[] GetResourceIds()
    {
        ThrowIfDisposed();
        var ptr = NativeMethods.undoc_get_resource_ids(_handle);
        if (ptr == IntPtr.Zero)
            return Array.Empty<string>();

        try
        {
            var json = PtrToStringUtf8(ptr);
            return JsonSerializer.Deserialize<string[]>(json) ?? Array.Empty<string>();
        }
        finally
        {
            NativeMethods.undoc_free_string(ptr);
        }
    }

    /// <summary>
    /// Get metadata for a resource.
    /// </summary>
    /// <param name="resourceId">The resource ID</param>
    /// <returns>Resource metadata as JSON, or null if not found</returns>
    public JsonDocument? GetResourceInfo(string resourceId)
    {
        ThrowIfDisposed();
        var ptr = NativeMethods.undoc_get_resource_info(_handle, resourceId);
        if (ptr == IntPtr.Zero)
            return null;

        try
        {
            var json = PtrToStringUtf8(ptr);
            return JsonDocument.Parse(json);
        }
        finally
        {
            NativeMethods.undoc_free_string(ptr);
        }
    }

    /// <summary>
    /// Get binary data for a resource.
    /// </summary>
    /// <param name="resourceId">The resource ID</param>
    /// <returns>Resource data as bytes, or null if not found</returns>
    public byte[]? GetResourceData(string resourceId)
    {
        ThrowIfDisposed();
        var ptr = NativeMethods.undoc_get_resource_data(_handle, resourceId, out var length);
        if (ptr == IntPtr.Zero)
            return null;

        try
        {
            var data = new byte[(int)length];
            Marshal.Copy(ptr, data, 0, data.Length);
            return data;
        }
        finally
        {
            NativeMethods.undoc_free_bytes(ptr, length);
        }
    }

    private static string GetLastError()
    {
        var ptr = NativeMethods.undoc_last_error();
        if (ptr == IntPtr.Zero)
            return "Unknown error";
        return Marshal.PtrToStringAnsi(ptr) ?? "Unknown error";
    }

    private static string PtrToStringUtf8(IntPtr ptr)
    {
        if (ptr == IntPtr.Zero)
            return string.Empty;

        // Find null terminator
        int len = 0;
        while (Marshal.ReadByte(ptr, len) != 0)
            len++;

        if (len == 0)
            return string.Empty;

        byte[] buffer = new byte[len];
        Marshal.Copy(ptr, buffer, 0, len);
        return Encoding.UTF8.GetString(buffer);
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(UndocDocument));
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (_handle != IntPtr.Zero)
            {
                NativeMethods.undoc_free_document(_handle);
                _handle = IntPtr.Zero;
            }
            _disposed = true;
        }
    }

    ~UndocDocument()
    {
        Dispose(false);
    }
}
