using System.Text.Json.Serialization;

namespace Undoc;

/// <summary>
/// One table of the document as delimited text — see <see cref="UndocDocument.GetTables"/>.
/// </summary>
public sealed class TableText
{
    /// <summary>
    /// The number of the section the table is in (from 1) — a sheet, a slide, a document
    /// section.
    /// </summary>
    [JsonPropertyName("section")]
    public int Section { get; init; }

    /// <summary>
    /// The table's place among that section's tables (from 1). A table nested in a cell
    /// comes right after the one that holds it.
    /// </summary>
    [JsonPropertyName("index")]
    public int Index { get; init; }

    /// <summary>
    /// The table as CSV (RFC 4180), or tab-separated text: a merged cell's text in its
    /// top-left position and the positions it covers empty, records ended with CRLF.
    /// </summary>
    [JsonPropertyName("text")]
    public string Text { get; init; } = "";
}
