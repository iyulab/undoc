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
    /// <summary>
    /// Why the call failed — lets a caller branch on the reason (report a damaged file
    /// on <see cref="UndocErrorKind.ZipArchive"/>, an unsupported input on
    /// <see cref="UndocErrorKind.UnsupportedFormat"/>) without matching on
    /// <see cref="Exception.Message"/>.
    /// </summary>
    /// <remarks>
    /// <see cref="UndocErrorKind.Other"/> when the failure did not come from the native
    /// library and so carries no classification. Never
    /// <see cref="UndocErrorKind.None"/>, which means success — a thrown exception is
    /// not a success.
    /// </remarks>
    public UndocErrorKind Kind { get; }

    /// <summary>
    /// Initialize an undoc exception with a native or wrapper error message.
    /// </summary>
    /// <param name="message">The error message.</param>
    public UndocException(string message) : this(message, UndocErrorKind.Other) { }

    /// <summary>
    /// Initialize an undoc exception with a message and the reason the call failed.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="kind">Why the call failed.</param>
    public UndocException(string message, UndocErrorKind kind) : base(message)
    {
        Kind = kind;
    }

    /// <summary>
    /// Initialize an undoc exception with a message and an inner exception.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="innerException">The underlying exception.</param>
    public UndocException(string message, Exception innerException) : base(message, innerException)
    {
        Kind = UndocErrorKind.Other;
    }
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

    /// <summary>
    /// Apply the lossless, idempotent markdown shape-refinement pass (table
    /// shape, ordered-list numbering, link/image paths, frontmatter, section
    /// anchors) after rendering. Default: <see langword="false"/>.
    /// </summary>
    public bool Refine { get; set; } = false;

    internal uint ToFlags()
    {
        uint flags = 0;
        if (IncludeFrontmatter) flags |= NativeMethods.UNDOC_FLAG_FRONTMATTER;
        if (EscapeSpecialChars) flags |= NativeMethods.UNDOC_FLAG_ESCAPE_SPECIAL;
        if (ParagraphSpacing) flags |= NativeMethods.UNDOC_FLAG_PARAGRAPH_SPACING;
        if (Refine) flags |= NativeMethods.UNDOC_FLAG_REFINE;
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
            return ptr == IntPtr.Zero ? "unknown" : PtrToStringUtf8(ptr);
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
            throw NativeFailure($"Failed to parse {path}");

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
                throw NativeFailure("Failed to parse bytes");

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
            throw NativeFailure("Failed to convert to markdown");

        return CopyAndFreeNativeUtf8String(ptr, NativeMethods.undoc_free_string);
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
            throw NativeFailure("Failed to convert to text");

        return CopyAndFreeNativeUtf8String(ptr, NativeMethods.undoc_free_string);
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
            throw NativeFailure("Failed to convert to JSON");

        return CopyAndFreeNativeUtf8String(ptr, NativeMethods.undoc_free_string);
    }

    /// <summary>
    /// Get plain text content (faster than ToText for simple extraction).
    /// </summary>
    /// <returns>Plain text string</returns>
    public string PlainText()
    {
        ThrowIfDisposed();
        var ptr = NativeMethods.undoc_plain_text(_handle);
        return CopyAndFreeRequiredNativeUtf8String(
            ptr,
            "Failed to get plain text",
            NativeFailure,
            NativeMethods.undoc_free_string);
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
                throw NativeFailure("Failed to get section count");
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
                throw NativeFailure("Failed to get resource count");
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

            return CopyAndFreeNativeUtf8String(ptr, NativeMethods.undoc_free_string);
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

            return CopyAndFreeNativeUtf8String(ptr, NativeMethods.undoc_free_string);
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
        return ParseResourceIdsFromNativeJson(ptr, NativeFailure, NativeMethods.undoc_free_string);
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

        var json = CopyAndFreeNativeUtf8String(ptr, NativeMethods.undoc_free_string);
        return JsonDocument.Parse(json);
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
        return PtrToStringUtf8(ptr);
    }

    /// <summary>
    /// Read the native classification of the last failure.
    /// </summary>
    /// <remarks>
    /// An unrecognised number is passed through unchanged rather than folded into
    /// <see cref="UndocErrorKind.Other"/>, so a newer native library stays usable and
    /// the raw value survives in logs. Zero is the one value that cannot stand: we are
    /// building a failure, and zero means success.
    /// </remarks>
    private static UndocErrorKind GetLastErrorKind()
    {
        var kind = NativeMethods.undoc_last_error_kind();
        return kind == (int)UndocErrorKind.None ? UndocErrorKind.Other : (UndocErrorKind)kind;
    }

    /// <summary>
    /// Build the exception for a failed native call, carrying both its message and its
    /// classification.
    /// </summary>
    /// <remarks>
    /// Every native failure goes through here so that no throw site can quietly drop
    /// the classification and leave the caller with <see cref="UndocErrorKind.Other"/>.
    /// </remarks>
    internal static UndocException NativeFailure(string operation) =>
        new UndocException($"{operation}: {GetLastError()}", GetLastErrorKind());

    internal static string CopyAndFreeNativeUtf8String(IntPtr ptr, Action<IntPtr> free)
    {
        if (ptr == IntPtr.Zero)
            return string.Empty;

        try
        {
            return PtrToStringUtf8(ptr);
        }
        finally
        {
            free(ptr);
        }
    }

    /// <remarks>
    /// Takes a factory that builds the whole exception rather than just its message, so
    /// the failure classification travels with it — a message-only seam would silently
    /// downgrade every failure routed through here to
    /// <see cref="UndocErrorKind.Other"/>.
    /// </remarks>
    internal static string CopyAndFreeRequiredNativeUtf8String(
        IntPtr ptr,
        string operation,
        Func<string, UndocException> nativeFailure,
        Action<IntPtr> free)
    {
        if (ptr == IntPtr.Zero)
            throw nativeFailure(operation);

        return CopyAndFreeNativeUtf8String(ptr, free);
    }

    internal static string[] ParseResourceIdsFromNativeJson(
        IntPtr ptr,
        Func<string, UndocException> nativeFailure,
        Action<IntPtr> free)
    {
        var json = CopyAndFreeRequiredNativeUtf8String(
            ptr,
            "Failed to get resource IDs",
            nativeFailure,
            free);
        return JsonSerializer.Deserialize<string[]>(json) ?? Array.Empty<string>();
    }

    internal static string PtrToStringUtf8(IntPtr ptr)
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

    /// <summary>
    /// Release the native document handle.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Release managed and unmanaged resources associated with this document.
    /// </summary>
    /// <param name="disposing">
    /// True when called from <see cref="Dispose()"/>; false when called from the finalizer.
    /// </param>
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

    /// <summary>
    /// Finalizer for releasing the native document handle if <see cref="Dispose()"/> was not called.
    /// </summary>
    ~UndocDocument()
    {
        Dispose(false);
    }
}
