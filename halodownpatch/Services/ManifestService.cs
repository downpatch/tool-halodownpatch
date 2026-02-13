using ExcelDataReader;
using halodownpatch.Models;

namespace halodownpatch.Services
{
    public sealed class ManifestService
    {
        private readonly HttpClient _http;

        public ManifestService(HttpClient http) => _http = http;

        public sealed class WorkbookData
        {
            public required List<ManifestRow> MccBase { get; init; }
            public required Dictionary<string, List<ManifestRow>> BySheet { get; init; }
            public required List<string> AvailableSheets { get; init; }
        }

        public async Task<WorkbookData> LoadWorkbookAsync(
            string xlsxPath = "data/ManifestDataSource.xlsx",
            CancellationToken ct = default)
        {
            var bytes = await _http.GetByteArrayAsync(xlsxPath, ct);
            using var ms = new MemoryStream(bytes);

            using var reader = ExcelReaderFactory.CreateReader(ms);

            var mccBase = new List<ManifestRow>();
            var bySheet = new Dictionary<string, List<ManifestRow>>(StringComparer.OrdinalIgnoreCase);
            var availableSheets = new List<string>();

            do
            {
                var sheetName = (reader.Name ?? "").Trim();
                if (string.IsNullOrWhiteSpace(sheetName))
                    continue;

                availableSheets.Add(sheetName);

                // Read header row
                if (!reader.Read())
                    continue;

                var headerIndex = BuildHeaderIndex(reader);

                var rows = new List<ManifestRow>();

                while (reader.Read())
                {
                    var slug = GetString(reader, headerIndex, "Slug");
                    if (string.IsNullOrWhiteSpace(slug))
                        continue;

                    var row = new ManifestRow
                    {
                        Sheet = sheetName,
                        Slug = slug.Trim(),
                        Name = GetString(reader, headerIndex, "Name").Trim(),
                        AppId = GetLong(reader, headerIndex, "AppID"),
                        DepotId = GetLong(reader, headerIndex, "DepotID"),
                        ManifestId = GetLong(reader, headerIndex, "ManifestID"),
                        TotalSizeBytes = GetNullableLong(reader, headerIndex, "TotalSizeBytes"),
                        ReleaseDateFull = GetString(reader, headerIndex, "ReleaseDateFull").Trim(),
                        MccRelease = GetString(reader, headerIndex, "MCC Release").Trim(),
                    };

                    rows.Add(row);
                }

                // Sort newest-first by suffix number when possible
                rows.Sort((a, b) => ExtractSlugNumber(b.Slug).CompareTo(ExtractSlugNumber(a.Slug)));

                if (sheetName.Equals("MCC Base", StringComparison.OrdinalIgnoreCase))
                {
                    mccBase = rows.Where(r => r.Name.Equals("MCC Base", StringComparison.OrdinalIgnoreCase)).ToList();
                }
                else
                {
                    bySheet[sheetName] = rows;
                }
            }
            while (reader.NextResult());

            if (mccBase.Count == 0)
                throw new InvalidOperationException("Sheet 'MCC Base' not found or contained no rows.");

            // Ensure base sorted too
            mccBase.Sort((a, b) => ExtractSlugNumber(b.Slug).CompareTo(ExtractSlugNumber(a.Slug)));

            return new WorkbookData
            {
                MccBase = mccBase,
                BySheet = bySheet,
                AvailableSheets = availableSheets
            };
        }

        public ManifestRow? FindMatchingRowForBase(ManifestRow selectedBase, List<ManifestRow> candidateSheetRows)
        {
            // Primary: match by MCC Release == base ReleaseDateFull
            var baseKey = NormalizeRelease(selectedBase.ReleaseDateFull);
            if (!string.IsNullOrWhiteSpace(baseKey))
            {
                var match = candidateSheetRows.FirstOrDefault(r =>
                    NormalizeRelease(r.MccRelease) == baseKey);

                if (match != null) return match;
            }

            // Fallback: match by suffix number (MCCBase_39 => Reach_39)
            var baseN = ExtractSlugNumber(selectedBase.Slug);
            if (baseN > 0)
            {
                var match = candidateSheetRows.FirstOrDefault(r => ExtractSlugNumber(r.Slug) == baseN);
                if (match != null) return match;
            }

            return null;
        }

        private static string NormalizeRelease(string s)
            => (s ?? "").Trim().Replace("\u00A0", " "); // handle odd spaces if present

        private static Dictionary<string, int> BuildHeaderIndex(IExcelDataReader reader)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                var h = reader.GetValue(i)?.ToString()?.Trim();
                if (!string.IsNullOrWhiteSpace(h) && !map.ContainsKey(h))
                    map[h] = i;
            }
            return map;
        }

        private static string GetString(IExcelDataReader reader, Dictionary<string, int> headerIndex, string col)
        {
            if (!headerIndex.TryGetValue(col, out var idx)) return "";
            return reader.GetValue(idx)?.ToString() ?? "";
        }

        private static long GetLong(IExcelDataReader reader, Dictionary<string, int> headerIndex, string col)
        {
            var s = GetString(reader, headerIndex, col).Trim().Trim('_');

            if (headerIndex.TryGetValue(col, out var idx))
            {
                var v = reader.GetValue(idx);
                if (v is double d) return checked((long)d);
                if (v is int i) return i;
                if (v is long l) return l;
                if (v is decimal m) return checked((long)m);
            }

            return long.TryParse(s, out var parsed) ? parsed : 0;
        }

        private static long? GetNullableLong(IExcelDataReader reader, Dictionary<string, int> headerIndex, string col)
        {
            var s = GetString(reader, headerIndex, col).Trim().Trim('_');

            if (headerIndex.TryGetValue(col, out var idx))
            {
                var v = reader.GetValue(idx);
                if (v is null) return null;
                if (v is double d) return checked((long)d);
                if (v is int i) return i;
                if (v is long l) return l;
                if (v is decimal m) return checked((long)m);
            }

            if (string.IsNullOrWhiteSpace(s)) return null;
            return long.TryParse(s, out var parsed) ? parsed : null;
        }

        private static int ExtractSlugNumber(string slug)
        {
            var idx = slug.LastIndexOf('_');
            if (idx >= 0 && idx + 1 < slug.Length && int.TryParse(slug[(idx + 1)..], out var n))
                return n;
            return 0;
        }
    }
}
