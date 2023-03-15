using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;

namespace ConsoleApp4.Excel
{
    public class PivotSettings
    {
        public List<PivotColumnLabel> ColumnLabels { get; set; } = new();
        public List<PivotRowLabel> RowLabels = new();
        public string[] FilterLabels { get; set; } = Array.Empty<string>();
        public string TableStyle { get; set; } = PivotBuiltInStyles.PivotStyleDark2;
        public string SheetName { get; set; } = "PIVOT";
        public string TableName { get; set; } = "PivotTable";

        public class PivotRowLabel
        {
            public string Name { get; set; }
            public ST_Axis Direction { get; set; }

            public ST_FieldSortType SortType { get; set; } = ST_FieldSortType.ascending;

            public bool Collapsed { get; set; } = true;
        }

        public class PivotColumnLabel
        {
            public string Name { get; set; }
            public DataConsolidateFunction DataConsolidateFunction { get; set; }
        };
    }

    public struct PivotBuiltInStyles
    {
        /* DARK STYLES */
        public const string PivotStyleDark1 = "PivotStyleDark1";
        public const string PivotStyleDark2 = "PivotStyleDark2";
        public const string PivotStyleDark3 = "PivotStyleDark3";
        public const string PivotStyleDark4 = "PivotStyleDark4";
        public const string PivotStyleDark5 = "PivotStyleDark5";
        public const string PivotStyleDark6 = "PivotStyleDark6";
        public const string PivotStyleDark7 = "PivotStyleDark7";
        public const string PivotStyleDark8 = "PivotStyleDark8";
        public const string PivotStyleDark9 = "PivotStyleDark9";
        public const string PivotStyleDark10 = "PivotStyleDark10";
        public const string PivotStyleDark11 = "PivotStyleDark11";
        public const string PivotStyleDark12 = "PivotStyleDark12";
        public const string PivotStyleDark13 = "PivotStyleDark13";
        public const string PivotStyleDark14 = "PivotStyleDark14";
        public const string PivotStyleDark15 = "PivotStyleDark15";
        public const string PivotStyleDark16 = "PivotStyleDark16";
        public const string PivotStyleDark17 = "PivotStyleDark17";
        public const string PivotStyleDark18 = "PivotStyleDark18";
        public const string PivotStyleDark19 = "PivotStyleDark19";
        public const string PivotStyleDark20 = "PivotStyleDark20";
        public const string PivotStyleDark21 = "PivotStyleDark21";
        public const string PivotStyleDark22 = "PivotStyleDark22";
        public const string PivotStyleDark23 = "PivotStyleDark23";
        public const string PivotStyleDark24 = "PivotStyleDark24";
        public const string PivotStyleDark25 = "PivotStyleDark25";
        public const string PivotStyleDark26 = "PivotStyleDark26";
        public const string PivotStyleDark27 = "PivotStyleDark27";
        public const string PivotStyleDark28 = "PivotStyleDark28";

        /* LIGHT STYLES */
        public const string PivotStyleLight1 = "PivotStyleLight1";
        public const string PivotStyleLight2 = "PivotStyleLight2";
        public const string PivotStyleLight3 = "PivotStyleLight3";
        public const string PivotStyleLight4 = "PivotStyleLight4";
        public const string PivotStyleLight5 = "PivotStyleLight5";
        public const string PivotStyleLight6 = "PivotStyleLight6";
        public const string PivotStyleLight7 = "PivotStyleLight7";
        public const string PivotStyleLight8 = "PivotStyleLight8";
        public const string PivotStyleLight9 = "PivotStyleLight9";
        public const string PivotStyleLight10 = "PivotStyleLight10";
        public const string PivotStyleLight11 = "PivotStyleLight11";
        public const string PivotStyleLight12 = "PivotStyleLight12";
        public const string PivotStyleLight13 = "PivotStyleLight13";
        public const string PivotStyleLight14 = "PivotStyleLight14";
        public const string PivotStyleLight15 = "PivotStyleLight15";
        public const string PivotStyleLight16 = "PivotStyleLight16";
        public const string PivotStyleLight17 = "PivotStyleLight17";
        public const string PivotStyleLight18 = "PivotStyleLight18";
        public const string PivotStyleLight19 = "PivotStyleLight19";
        public const string PivotStyleLight20 = "PivotStyleLight20";
        public const string PivotStyleLight21 = "PivotStyleLight21";
        public const string PivotStyleLight22 = "PivotStyleLight22";
        public const string PivotStyleLight23 = "PivotStyleLight23";
        public const string PivotStyleLight24 = "PivotStyleLight24";
        public const string PivotStyleLight25 = "PivotStyleLight25";
        public const string PivotStyleLight26 = "PivotStyleLight26";
        public const string PivotStyleLight27 = "PivotStyleLight27";
        public const string PivotStyleLight28 = "PivotStyleLight28";

        /* MEDIUM STYLES */
        public const string PivotStyleMedium1 = "PivotStyleMedium1";
        public const string PivotStyleMedium2 = "PivotStyleMedium2";
        public const string PivotStyleMedium3 = "PivotStyleMedium3";
        public const string PivotStyleMedium4 = "PivotStyleMedium4";
        public const string PivotStyleMedium5 = "PivotStyleMedium5";
        public const string PivotStyleMedium6 = "PivotStyleMedium6";
        public const string PivotStyleMedium7 = "PivotStyleMedium7";
        public const string PivotStyleMedium8 = "PivotStyleMedium8";
        public const string PivotStyleMedium9 = "PivotStyleMedium9";
        public const string PivotStyleMedium10 = "PivotStyleMedium10";
        public const string PivotStyleMedium11 = "PivotStyleMedium11";
        public const string PivotStyleMedium12 = "PivotStyleMedium12";
        public const string PivotStyleMedium13 = "PivotStyleMedium13";
        public const string PivotStyleMedium14 = "PivotStyleMedium14";
        public const string PivotStyleMedium15 = "PivotStyleMedium15";
        public const string PivotStyleMedium16 = "PivotStyleMedium16";
        public const string PivotStyleMedium17 = "PivotStyleMedium17";
        public const string PivotStyleMedium18 = "PivotStyleMedium18";
        public const string PivotStyleMedium19 = "PivotStyleMedium19";
        public const string PivotStyleMedium20 = "PivotStyleMedium20";
        public const string PivotStyleMedium21 = "PivotStyleMedium21";
        public const string PivotStyleMedium22 = "PivotStyleMedium22";
        public const string PivotStyleMedium23 = "PivotStyleMedium23";
        public const string PivotStyleMedium24 = "PivotStyleMedium24";
        public const string PivotStyleMedium25 = "PivotStyleMedium25";
        public const string PivotStyleMedium26 = "PivotStyleMedium26";
        public const string PivotStyleMedium27 = "PivotStyleMedium27";
        public const string PivotStyleMedium28 = "PivotStyleMedium28";


    }
}
