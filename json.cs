// Tabular Editor Advanced Scripting (C#)
// Export schema for LLM use: tables, columns, measures (with DAX), relationships.

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;

string outPath = @"C:\temp\semantic_model.json"; // <-- CHANGE THIS

Func<string, string> esc = s =>
    (s ?? "")
    .Replace("\\", "\\\\")
    .Replace("\"", "\\\"")
    .Replace("\r", " ")
    .Replace("\n", " ");

var tables = Model.Tables
    .Where(t => !t.IsHidden)
    .Select(t => new {
        name = t.Name,
        columns = t.Columns
            .Where(c => !c.IsHidden)
            .Select(c => new {
                name = c.Name,
                dataType = c.DataType.ToString()
            })
            .ToList(),
        measures = t.Measures
            .Where(m => !m.IsHidden)
            .Select(m => new {
                name = m.Name,
                definition = m.Expression
            })
            .ToList()
    })
    .ToList();

var relationships = Model.Relationships
    .Select(r => new {
        fromTable = r.FromTable.Name,
        fromColumn = r.FromColumn.Name,
        toTable = r.ToTable.Name,
        toColumn = r.ToColumn.Name,
        isActive = r.IsActive
    })
    .ToList();

// Date column hints (helps your default YoY logic)
var dateCandidates = Model.Tables
    .SelectMany(t => t.Columns.Select(c => new { table = t.Name, col = c }))
    .Where(x => !x.col.IsHidden)
    .Where(x =>
        x.col.DataType.ToString().Equals("DateTime", StringComparison.OrdinalIgnoreCase) ||
        x.col.Name.Equals("Date", StringComparison.OrdinalIgnoreCase) ||
        x.col.Name.EndsWith("Date", StringComparison.OrdinalIgnoreCase)
    )
    .Select(x => $"{x.table}[{x.col.Name}]")
    .Distinct()
    .Take(100)
    .ToList();

var sb = new StringBuilder();
sb.Append("{\n");

sb.Append("  \"tables\": [\n");
for (int i = 0; i < tables.Count; i++)
{
    var t = tables[i];
    sb.Append("    {\n");
    sb.Append($"      \"name\": \"{esc(t.name)}\",\n");

    sb.Append("      \"columns\": [\n");
    for (int j = 0; j < t.columns.Count; j++)
    {
        var c = t.columns[j];
        sb.Append("        {\n");
        sb.Append($"          \"name\": \"{esc(c.name)}\",\n");
        sb.Append($"          \"dataType\": \"{esc(c.dataType)}\"\n");
        sb.Append("        }");
        sb.Append(j < t.columns.Count - 1 ? ",\n" : "\n");
    }
    sb.Append("      ],\n");

    sb.Append("      \"measures\": [\n");
    for (int j = 0; j < t.measures.Count; j++)
    {
        var m = t.measures[j];
        sb.Append("        {\n");
        sb.Append($"          \"name\": \"{esc(m.name)}\",\n");
        sb.Append($"          \"definition\": \"{esc(m.definition)}\"\n");
        sb.Append("        }");
        sb.Append(j < t.measures.Count - 1 ? ",\n" : "\n");
    }
    sb.Append("      ]\n");

    sb.Append("    }");
    sb.Append(i < tables.Count - 1 ? ",\n" : "\n");
}
sb.Append("  ],\n");

sb.Append("  \"relationships\": [\n");
for (int i = 0; i < relationships.Count; i++)
{
    var r = relationships[i];
    sb.Append("    {\n");
    sb.Append($"      \"from\": \"{esc(r.fromTable)}[{esc(r.fromColumn)}]\",\n");
    sb.Append($"      \"to\": \"{esc(r.toTable)}[{esc(r.toColumn)}]\",\n");
    sb.Append($"      \"active\": {r.isActive.ToString().ToLower()}\n");
    sb.Append("    }");
    sb.Append(i < relationships.Count - 1 ? ",\n" : "\n");
}
sb.Append("  ],\n");

sb.Append("  \"hints\": {\n");
sb.Append("    \"date_column_candidates\": [");
sb.Append(string.Join(", ", dateCandidates.Select(x => $"\"{esc(x)}\"")));
sb.Append("]\n");
sb.Append("  },\n");

sb.Append("  \"rules\": [\n");
sb.Append("    \"Use only names found in this schema; do not invent tables/columns/measures.\",\n");
sb.Append("    \"Generate one DAX query starting with EVALUATE that returns exactly one table.\",\n");
sb.Append("    \"Keep results <= 200 rows unless the user asks otherwise.\",\n");
sb.Append("    \"If YoY is requested and period not specified, default to YTD vs YTD last year using the best Date column.\"\n");
sb.Append("  ]\n");

sb.Append("}\n");

Directory.CreateDirectory(Path.GetDirectoryName(outPath));
File.WriteAllText(outPath, sb.ToString(), Encoding.UTF8);

Console.WriteLine(\"Wrote schema to: \" + outPath);
