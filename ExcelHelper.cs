using FastMember;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ConsoleApp4.Excel
{
    public static class ExcelHelper
    {
        private static readonly Dictionary<Type, Action<ICell, object>> cellFormatMapper = new()
        {
            { typeof(int), (ICell cell, object value) =>  cell.SetCellValue((int) value) },
            { typeof(double), (ICell cell, object value) => cell.SetCellValue((double) value) },
            { typeof(DateTime), (ICell cell, object value) => cell.SetCellValue((DateTime) value)},
            { typeof(string), (ICell cell, object value) => cell.SetCellValue((string) value) },
            { typeof(decimal), (ICell cell, object value) => cell.SetCellValue(decimal.ToDouble((decimal) value)) },
            { typeof(bool), (ICell cell, object value) => cell.SetCellValue((bool) value) },
            { typeof(byte), (ICell cell, object value) => cell.SetCellValue(Convert.ToDouble((byte) value)) },
            { typeof(ushort), (ICell cell, object value) => cell.SetCellValue(Convert.ToDouble((ushort) value)) },
            { typeof(short), (ICell cell, object value) => cell.SetCellValue(Convert.ToDouble((short) value)) },
            { typeof(long), (ICell cell, object value) => cell.SetCellValue(Convert.ToDouble((long) value)) },
            { typeof(char), (ICell cell, object value) => cell.SetCellValue(((char) value).ToString()) },
            { typeof(float), (ICell cell, object value) => cell.SetCellValue(Convert.ToDouble((float) value)) },
            { typeof(Guid), (ICell cell, object value) => cell.SetCellValue(((Guid) value).ToString()) },
            { typeof(object), (ICell cell, object value) => cell.SetCellValue(value.ToString()) }
        };

        /// <summary>
        /// Writes an <see cref="IEnumerable{T}"/> into a stream, 
        /// optionally creates a pivot table based on produced table
        /// if the <see cref="PivotSettings"/> is not null
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <param name="objects"></param>
        /// <param name="visibleColumns"></param>
        /// <param name="settings"></param>
        /// <param name="leaveOpen"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void WriteToExcelTable<T>(Stream stream, IEnumerable<T> objects, string[] ignoredColumns = null, PivotSettings pivotSettings = null, bool leaveOpen = false)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (objects == null) throw new ArgumentNullException(nameof(objects));

            // Uses object reader to itterate over the columns of the records
            using var reader = ObjectReader.Create(objects);
            List<TableColumn> columns = GetVisibleColumns(reader, ignoredColumns).ToList();

            if (!columns.Any())
                throw new InvalidOperationException("There are no writable columns");

            UpdateColumnWidths(columns);

            // Create Worksheet
            using IWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("DATA");

            // Create table
            XSSFTable table = CreateTable(sheet, columns);
            int rowCount = FillData(sheet, reader, columns);

            // Manually auto size columns, is way faster than sheet's AutoSizeColumn() method
            ResizeColumns(sheet, columns);

            // Update table's referenced area
            // see bug: https://github.com/nissl-lab/npoi/issues/1026
            // also see: https://github.com/nissl-lab/npoi/pull/1035
            AreaReference dataRange = new(new CellReference(0, 0), new CellReference(rowCount - 1, columns.Count - 1));
            table.SetCellReferences(dataRange);

            // Add column filters
            table.GetCTTable().autoFilter = new()
            {
                @ref = dataRange.FormatAsString()
            };

            // Create pivot table
            if (pivotSettings != null)
            {
                XSSFPivotTable pivotTable = PivotHelper.CreatePivotTable(table, pivotSettings);
                sheet.IsSelected = false;
                workbook.SetActiveSheet(workbook.GetSheetIndex(pivotTable.GetParentSheet().SheetName));
            }

            // write book to the stream
            workbook.Write(stream, leaveOpen);
        }

        /// <summary>
        /// Initializes the <see cref="XSSFTable"/>
        /// with default style and columns.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columns"></param>
        /// <returns>An instance of <see cref="XSSFTable"/></returns>
        private static XSSFTable CreateTable(XSSFSheet sheet, IList<TableColumn> columns)
        {
            XSSFTable xssfTable = sheet.CreateTable();

            xssfTable.GetCTTable().id = 1;
            xssfTable.Name = "Data";
            xssfTable.IsHasTotalsRow = false;
            xssfTable.DisplayName = "MYTABLE";

            // see bug: https://github.com/nissl-lab/npoi/issues/1026
            // also see: https://github.com/nissl-lab/npoi/pull/1035
            xssfTable.SetCellReferences(new AreaReference(new CellReference(0, 0), new CellReference(1, 1), SpreadsheetVersion.EXCEL2007));

            xssfTable.StyleName = XSSFBuiltinTableStyleEnum.TableStyleMedium16.ToString();
            xssfTable.Style.IsShowColumnStripes = false;
            xssfTable.Style.IsShowRowStripes = true;

            for (int i = 0; i < columns.Count; i++)
            {
                xssfTable.CreateColumn(columns[i].Name, i);
            }

            return xssfTable;
        }


        /// <summary>
        /// Fills table data and returns the number of created rows
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="reader"></param>
        /// <param name="columns"></param>
        /// <param name="colWidths"></param>
        /// <returns></returns>
        private static int FillData(XSSFSheet sheet, ObjectReader reader, IList<TableColumn> columns)
        {
            // fill the header row
            IRow headerRow = sheet.CreateRow(0);

            foreach (TableColumn c in columns)
            {
                var cell = headerRow.CreateCell(c.Index);
                cell.SetCellValue(c.Name);
            }

            // populate values
            int index = 1;
            object[] values = new object[columns.Count];

            while (reader.Read())
            {
                var row = sheet.CreateRow(index);
                int instances = reader.GetValues(values);

                for (int j = 0; j < instances; j++)
                {
                    ICell cell = row.CreateCell(j);

                    // invoke setvalue method with respective type based on type of the boxed object
                    object value = values[j];
                    cellFormatMapper[value.GetType()].Invoke(cell, value);

                    //compare lengths
                    int length = value.ToString().Length;

                    if (columns[j].Width < length)
                    {
                        columns[j].Width = length + 2;
                    }
                }
                index++;
            }

            return index;
        }

        private static IEnumerable<TableColumn> GetVisibleColumns(ObjectReader reader, string[] ignoredColumns = null)
        {
            string[] readerColumns = new string[reader.FieldCount];

            for (int i = 0; i < reader.FieldCount; i++)
            {
                readerColumns[i] = reader.GetName(i);
            }

            ignoredColumns ??= Array.Empty<string>();

            IEnumerable<string> visibleColumns = readerColumns.Except(ignoredColumns);

            int index = 0;

            foreach (var col in visibleColumns)
            {
                yield return new TableColumn()
                {
                    Name = col,
                    Index = index++,
                    Width = 0
                };
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="columns"></param>
        /// <param name="resArray"></param>
        private static void UpdateColumnWidths(IEnumerable<TableColumn> columns)
        {
            foreach(TableColumn c in columns)
            {
                int length = c.Name.Length;

                if (c.Width < length)
                    c.Width = length + 4;
            }
        }

        /// <summary>
        /// Resizes columns and sets precalculated width
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columns"></param>
        /// <param name="colWidths"></param>
        private static void ResizeColumns(XSSFSheet sheet, IList<TableColumn> columns)
        {
            for(int i=0; i<columns.Count; i++)
            {
                int width = (int) (columns[i].Width  * 1.25f) * 256; // floating value can be adjusted at any time
                sheet.SetColumnWidth(i, Math.Min(width, 65279));
            }
        }
    }

    internal class TableColumn
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public int Width { get; set; }
    }
}