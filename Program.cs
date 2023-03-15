using System.Configuration;
using ConsoleApp4.Excel;
using FastMember;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using QueryManager;
using static ConsoleApp4.Excel.PivotSettings;

namespace ConsoleApp
{

    internal class Program
    {
        private static readonly string ConnectionString = ConfigurationManager.ConnectionStrings["Default"].ConnectionString;

        private static void Main()
        {
            var pivotSettings = new PivotSettings()
            {
                ColumnLabels = new List<PivotColumnLabel>()
                {
                    new PivotColumnLabel { Name = "CA_NET", DataConsolidateFunction = DataConsolidateFunction.SUM },
                    new PivotColumnLabel { Name = "CA_BRUT", DataConsolidateFunction = DataConsolidateFunction.SUM },
                    new PivotColumnLabel { Name = "QTE_VENDUE", DataConsolidateFunction = DataConsolidateFunction.SUM },
                    new PivotColumnLabel { Name = "REMISE_MOYENNE", DataConsolidateFunction = DataConsolidateFunction.AVERAGE },
                },
                RowLabels = new List<PivotRowLabel>(){
                    new PivotRowLabel { Name = "GAMME", Direction = ST_Axis.axisRow },
                    new PivotRowLabel { Name = "SOUS_GAMME", Direction = ST_Axis.axisRow },
                    new PivotRowLabel { Name = "PRODUIT", Direction = ST_Axis.axisRow },
                    new PivotRowLabel { Name = "PERIODE", Direction = ST_Axis.axisCol },
                    new PivotRowLabel { Name = "MOIS", Direction = ST_Axis.axisCol },
                }
            };

            string[] ignoredColumns = new[] { "TEST" };
            using var stream = File.Open("test.xlsx", FileMode.Create, FileAccess.Write);

            ExcelHelper.WriteToExcelTable(stream, ReadFromStockProc(), null, pivotSettings, false);
        }


        private static IEnumerable<StatPharmacie> ReadFromStockProc()
        {
            string command = "StatistiquePharmacie";

            object parameters = new
            {
                @PHARMACIE = null as string,
                @LABORATOIRE = null as string,
                @GAMME = "CAP",
                @SOUS_GAMME = null as string,
                @PERIODE_UN_DEBUT = new DateTime(2022, 11, 01),
                @PERIODE_UN_FIN = new DateTime(2023, 12, 31),
                @PERIODE_DEUX_DEBUT = new DateTime(2022, 09, 01),
                @PERIODE_DEUX_FIN = new DateTime(2022, 10, 31),
                @LABO_GROUPED = true,
                @GAMME_GROUPED = true,
                @SOUS_GAMME_GROUPED = true,
                @PRODUIT_GROUPED = true,
                @MOIS_GROUPED = true
            };

            foreach (var item in DbManager.Read<StatPharmacie>(ConnectionString, command, System.Data.CommandType.StoredProcedure, parameters))
            {
                yield return item;
            }
        }
    }

    public class StatPharmacie
    {
        [Ordinal(0)]
        public string NOM { get; set; }
        [Ordinal(1)]
        public int CODE_IVRYLAB { get; set; }
        [Ordinal(2)]
        public int CIPGERS { get; set; }
        [Ordinal(3)]
        public string VILLE { get; set; }
        [Ordinal(4)]
        public string CP { get; set; }
        [Ordinal(5)]
        public string GAMME { get; set; }
        [Ordinal(6)]
        public string SOUS_GAMME { get; set; }
        [Ordinal(7)]
        public string FABRIQUANT { get; set; }
        [Ordinal(8)]
        public string PRODUIT { get; set; }
        [Ordinal(9)]
        public string MOIS { get; set; }
        [Ordinal(10)]
        public string PERIODE { get; set; }
        [Ordinal(11)]
        public decimal CA_NET { get; set; }
        [Ordinal(12)]
        public decimal CA_BRUT { get; set; }
        [Ordinal(13)]
        public decimal REMISE_MOYENNE { get; set; }
        [Ordinal(14)]
        public int QTE_VENDUE { get; set; }
    }

}


