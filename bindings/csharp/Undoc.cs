/*
 * undoc - C# Bindings
 *
 * P/Invoke wrapper for the undoc library.
 *
 * Usage:
 *   using Iyulab.Undoc;
 *
 *   using var doc = UndocDocument.FromFile("document.docx");
 *   string markdown = doc.ToMarkdown();
 *   Console.WriteLine(markdown);
 *
 * Copyright (c) 2024 iyulab
 * MIT License
 */

using System;
using System.Runtime.InteropServices;

namespace Iyulab.Undoc
{
    /// <summary>
    /// Flags for markdown rendering.
    /// </summary>
    [Flags]
    public enum MarkdownFlags
    {
        /// <summary>No flags.</summary>
        None = 0,
        /// <summary>Include YAML frontmatter with metadata.</summary>
        Frontmatter = 1,
        /// <summary>Escape special Markdown characters.</summary>
        EscapeSpecial = 2,
        /// <summary>Add blank lines between paragraphs.</summary>
        ParagraphSpacing = 4
    }

    /// <summary>
    /// JSON format options.
    /// </summary>
    public enum JsonFormat
    {
        /// <summary>Pretty-printed JSON with indentation.</summary>
        Pretty = 0,
        /// <summary>Compact JSON without whitespace.</summary>
        Compact = 1
    }

    /// <summary>
    /// Native interop methods.
    /// </summary>
    internal static class UndocNative
    {
        private const string LibraryName = "undoc";

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr undoc_version();

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr undoc_last_error();

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        public static extern IntPtr undoc_parse_file([MarshalAs(UnmanagedType.LPUTF8Str)] string path);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr undoc_parse_bytes(byte[] data, UIntPtr len);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void undoc_free_document(IntPtr doc);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr undoc_to_markdown(IntPtr doc, int flags);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr undoc_to_text(IntPtr doc);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr undoc_to_json(IntPtr doc, int format);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr undoc_plain_text(IntPtr doc);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int undoc_section_count(IntPtr doc);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int undoc_resource_count(IntPtr doc);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr undoc_get_title(IntPtr doc);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr undoc_get_author(IntPtr doc);

        [DllImport(LibraryName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void undoc_free_string(IntPtr str);
    }

    /// <summary>
    /// Exception thrown when undoc operations fail.
    /// </summary>
    public class UndocException : Exception
    {
        public UndocException(string message) : base(message) { }
    }

    /// <summary>
    /// Represents a parsed Office document.
    /// </summary>
    public class UndocDocument : IDisposable
    {
        private IntPtr _handle;
        private bool _disposed;

        /// <summary>
        /// Get the library version.
        /// </summary>
        public static string Version
        {
            get
            {
                var ptr = UndocNative.undoc_version();
                return Marshal.PtrToStringAnsi(ptr) ?? "";
            }
        }

        private UndocDocument(IntPtr handle)
        {
            _handle = handle;
        }

        /// <summary>
        /// Parse a document from a file path.
        /// </summary>
        /// <param name="path">Path to the document file.</param>
        /// <returns>Parsed document.</returns>
        /// <exception cref="UndocException">If parsing fails.</exception>
        public static UndocDocument FromFile(string path)
        {
            var handle = UndocNative.undoc_parse_file(path);
            if (handle == IntPtr.Zero)
            {
                throw new UndocException(GetLastError());
            }
            return new UndocDocument(handle);
        }

        /// <summary>
        /// Parse a document from a byte array.
        /// </summary>
        /// <param name="data">Document data.</param>
        /// <returns>Parsed document.</returns>
        /// <exception cref="UndocException">If parsing fails.</exception>
        public static UndocDocument FromBytes(byte[] data)
        {
            var handle = UndocNative.undoc_parse_bytes(data, (UIntPtr)data.Length);
            if (handle == IntPtr.Zero)
            {
                throw new UndocException(GetLastError());
            }
            return new UndocDocument(handle);
        }

        /// <summary>
        /// Convert the document to Markdown.
        /// </summary>
        /// <param name="flags">Rendering flags.</param>
        /// <returns>Markdown string.</returns>
        public string ToMarkdown(MarkdownFlags flags = MarkdownFlags.None)
        {
            ThrowIfDisposed();
            var ptr = UndocNative.undoc_to_markdown(_handle, (int)flags);
            if (ptr == IntPtr.Zero)
            {
                throw new UndocException(GetLastError());
            }
            try
            {
                return Marshal.PtrToStringUTF8(ptr) ?? "";
            }
            finally
            {
                UndocNative.undoc_free_string(ptr);
            }
        }

        /// <summary>
        /// Convert the document to plain text.
        /// </summary>
        /// <returns>Plain text string.</returns>
        public string ToText()
        {
            ThrowIfDisposed();
            var ptr = UndocNative.undoc_to_text(_handle);
            if (ptr == IntPtr.Zero)
            {
                throw new UndocException(GetLastError());
            }
            try
            {
                return Marshal.PtrToStringUTF8(ptr) ?? "";
            }
            finally
            {
                UndocNative.undoc_free_string(ptr);
            }
        }

        /// <summary>
        /// Convert the document to JSON.
        /// </summary>
        /// <param name="format">JSON format.</param>
        /// <returns>JSON string.</returns>
        public string ToJson(JsonFormat format = JsonFormat.Pretty)
        {
            ThrowIfDisposed();
            var ptr = UndocNative.undoc_to_json(_handle, (int)format);
            if (ptr == IntPtr.Zero)
            {
                throw new UndocException(GetLastError());
            }
            try
            {
                return Marshal.PtrToStringUTF8(ptr) ?? "";
            }
            finally
            {
                UndocNative.undoc_free_string(ptr);
            }
        }

        /// <summary>
        /// Get the plain text content.
        /// </summary>
        public string PlainText
        {
            get
            {
                ThrowIfDisposed();
                var ptr = UndocNative.undoc_plain_text(_handle);
                if (ptr == IntPtr.Zero)
                {
                    return "";
                }
                try
                {
                    return Marshal.PtrToStringUTF8(ptr) ?? "";
                }
                finally
                {
                    UndocNative.undoc_free_string(ptr);
                }
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
                var count = UndocNative.undoc_section_count(_handle);
                if (count < 0)
                {
                    throw new UndocException(GetLastError());
                }
                return count;
            }
        }

        /// <summary>
        /// Get the number of embedded resources.
        /// </summary>
        public int ResourceCount
        {
            get
            {
                ThrowIfDisposed();
                var count = UndocNative.undoc_resource_count(_handle);
                if (count < 0)
                {
                    throw new UndocException(GetLastError());
                }
                return count;
            }
        }

        /// <summary>
        /// Get the document title.
        /// </summary>
        public string? Title
        {
            get
            {
                ThrowIfDisposed();
                var ptr = UndocNative.undoc_get_title(_handle);
                if (ptr == IntPtr.Zero)
                {
                    return null;
                }
                try
                {
                    return Marshal.PtrToStringUTF8(ptr);
                }
                finally
                {
                    UndocNative.undoc_free_string(ptr);
                }
            }
        }

        /// <summary>
        /// Get the document author.
        /// </summary>
        public string? Author
        {
            get
            {
                ThrowIfDisposed();
                var ptr = UndocNative.undoc_get_author(_handle);
                if (ptr == IntPtr.Zero)
                {
                    return null;
                }
                try
                {
                    return Marshal.PtrToStringUTF8(ptr);
                }
                finally
                {
                    UndocNative.undoc_free_string(ptr);
                }
            }
        }

        private static string GetLastError()
        {
            var ptr = UndocNative.undoc_last_error();
            if (ptr == IntPtr.Zero)
            {
                return "Unknown error";
            }
            return Marshal.PtrToStringAnsi(ptr) ?? "Unknown error";
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(UndocDocument));
            }
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
                    UndocNative.undoc_free_document(_handle);
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
}
