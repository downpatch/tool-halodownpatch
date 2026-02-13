namespace halodownpatch.Models
{
    public sealed class ManifestRow
    {
        public string Sheet { get; set; } = "";
        public string Slug { get; set; } = "";
        public string Name { get; set; } = "";
        public long AppId { get; set; }
        public long DepotId { get; set; }
        public long ManifestId { get; set; }
        public long? TotalSizeBytes { get; set; }
        public string ReleaseDateFull { get; set; } = "";

        public string MccRelease { get; set; } = "";
    }
}
