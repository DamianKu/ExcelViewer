using System;
using System.Collections.Generic;
using System.Linq;
using ExcelViewerV2.Model;

namespace ExcelViewerV2.Search
{
    public sealed class SpreadsheetSearchService
    {
        private readonly IReadOnlyList<SpreadsheetModel> _sheets;

        public SpreadsheetSearchService(IReadOnlyList<SpreadsheetModel> sheets) => _sheets = sheets;

public List<SearchResult> FindAll(string query, bool matchCase, bool matchWholeCell)
{
    var results = new List<SearchResult>();
    if (string.IsNullOrWhiteSpace(query))
        return results;

    var comparison = matchCase
        ? StringComparison.Ordinal
        : StringComparison.OrdinalIgnoreCase;

    foreach (var sheet in _sheets)
    {
        if (!sheet.IsVisible) continue; // 🔹 pomijamy ukryte arkusze

        for (int r = 0; r < sheet.Rows; r++)
        {
            for (int c = 0; c < sheet.Columns; c++)
            {
                var cell = sheet[r, c];
                if (string.IsNullOrEmpty(cell.Text)) continue;

                bool match = matchWholeCell
                    ? string.Equals(cell.Text, query, comparison)
                    : cell.Text.IndexOf(query, comparison) >= 0;

                if (match)
                    results.Add(new SearchResult(sheet.Name, r, c, cell.Text));
            }
        }
    }

    return results;
}
    }
}
