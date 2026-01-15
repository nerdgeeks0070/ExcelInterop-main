

namespace Spreadsheet.Handler.Objects
{
    internal class DissolvedComponentResult
    {
        public string ResultSetId { get; set; }

        public string ResultId { get; set; }

        public string PeakName { get; set; }

        public string ResultSetIdPeakNameKey { get { return ResultSetId + ":" + PeakName; } }

        public string Bath { get; set; }

        public string RoundedDissolvedAmount { get; set; }

        public double TransferTime { get; set; }

        public string Vessel { get; set; }

    }
}
