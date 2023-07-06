using FastMember;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ConsoleApp4.Excel
{
   internal readonly struct DataFormats
    {
        public const string TEXT = "@";
        public const string GENERAL = "General";
        public const string INTEGER = "#,##0";
        public const string FLOATING = "#,##0.00";
        public const string DATE = "d-mmm-yy";
        public const string DATETIME = "d-mmm-yy h:mm:ss";
    }

    public class ExcelHelper
    {
        private readonly Dictionary<Type, Action<ICell, IDataFormat, object>> _cellValueSetMapper = new()
        {
            { typeof(int), (c, df, v) =>  SetValueAndFormat(c, df, (int) v) },
            { typeof(double), (c, df, v) => SetValueAndFormat(c, df, (double) v) },
            { typeof(DateTime), (c, df, v) => SetValueAndFormat(c, df, (DateTime) v)},
            { typeof(string), (c, df, v) => SetValueAndFormat(c, df, (string) v) },
            { typeof(decimal), (c, df, v) => SetValueAndFormat(c, df, decimal.ToDouble((decimal) v)) },
            { typeof(bool), (c, df, v) => SetValueAndFormat(c, df, (bool) v) },
            { typeof(byte), (c, df, v) => SetValueAndFormat(c, df, (byte) v) },
            { typeof(short), (c, df, v) => SetValueAndFormat(c, df, (short) v) },
            { typeof(long), (c, df, v) => SetValueAndFormat(c, df, Convert.ToDouble((long) v)) },
            { typeof(char), (c, df, v) => SetValueAndFormat(c, df, ((char) v).ToString()) },
            { typeof(float), (c, df, v) => SetValueAndFormat(c, df, (float) v) },
            { typeof(Guid), (c, df, v) => SetValueAndFormat(c, df, ((Guid) v).ToString()) },
            { typeof(object), (c, df, v) => SetValueAndFormat(c, df, v.ToString()) }
        };

        private static Dictionary<Type, ICellStyle> _cellStyleCache;

        private static readonly IDictionary<Type, string> _cellFormatMapper = new Dictionary<Type, string>() {
            { typeof(double), DataFormats.FLOATING},
            { typeof(byte), DataFormats.INTEGER },
            { typeof(int), DataFormats.INTEGER },
            { typeof(float), DataFormats.FLOATING },
            { typeof(decimal), DataFormats.FLOATING },
            { typeof(short), DataFormats.INTEGER },
            { typeof(long), DataFormats.INTEGER },
            { typeof(char), DataFormats.TEXT },
            { typeof(string), DataFormats.TEXT },
            { typeof(Guid), DataFormats.TEXT },
            { typeof(DateTime), DataFormats.DATE },
            { typeof(object), DataFormats.TEXT },
        };

        public static IDictionary<Type, string> DataFormatMapping { get; set; }

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
        public void Write<T>(Stream stream, IEnumerable<T> objects, bool leaveOpen = false, params string[] members)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (objects == null) throw new ArgumentNullException(nameof(objects));

            // Uses object reader to itterate over the columns of the records
            using var reader = ObjectReader.Create(objects, members);
            IList<TableColumn> columns = GetVisibleColumns(reader).ToList();

            if (columns.Count == 0)
                throw new InvalidOperationException("There are no writable columns");

            // Create Worksheet
            using IWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("DATA");

            _cellStyleCache = new();

            // Write data
            int rowCount = FillData(sheet, reader, columns);

            // Create table
            XSSFTable table = CreateTable(sheet, columns, rowCount);

            // Manually auto size columns, is way faster than sheet's AutoSizeColumn() method
            ResizeColumns(sheet, columns);

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
        private static XSSFTable CreateTable(XSSFSheet sheet, IList<TableColumn> columns, int rowCount)
        {
            XSSFTable xssfTable = sheet.CreateTable();

            xssfTable.GetCTTable().id = 1;
            xssfTable.Name = "Data";
            xssfTable.IsHasTotalsRow = false;
            xssfTable.DisplayName = "MYTABLE";

            var tableRange = new AreaReference(new CellReference(0, 0), new CellReference(rowCount - 1, columns.Count - 1));
            xssfTable.SetCellReferences(tableRange);

            xssfTable.StyleName = XSSFBuiltinTableStyleEnum.TableStyleMedium16.ToString();
            xssfTable.Style.IsShowColumnStripes = false;
            xssfTable.Style.IsShowRowStripes = true;

            for (int i = 0; i < columns.Count; i++)
            {
                xssfTable.CreateColumn(columns[i].Name, i);
            }

            // Add column filters
            xssfTable.GetCTTable().autoFilter = new()
            {
                @ref = tableRange.FormatAsString()
            };

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
        private int FillData(XSSFSheet sheet, ObjectReader reader, IList<TableColumn> columns)
        {
            IDataFormat dataFormat = sheet.Workbook.CreateDataFormat();

            // fill the header row
            IRow headerRow = sheet.CreateRow(0);

            for (int i=0; i < columns.Count; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(columns[i].Name);
            }

            // populate values
            int rowCount = 1;

            while (reader.Read())
            {
                var row = sheet.CreateRow(rowCount);

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    ICell cell = row.CreateCell(i);

                    // invoke setvalue method with respective type based on type of the boxed object
                    object value = reader.GetValue(i);
                    var setCellValueFunc = _cellValueSetMapper[value.GetType()];
                    setCellValueFunc(cell, dataFormat, value);

                    //compare lengths
                    int length = value.ToString().Length;

                    if (columns[i].Width < length)
                    {
                        columns[i].Width = length + 2;
                    }
                }
                rowCount++;
            }

            return rowCount;
        }

        private static IList<TableColumn> GetVisibleColumns(ObjectReader reader, params string[] members)
        {
            var tableColumns = new List<TableColumn>();

            if (reader.FieldCount == 0)
                return tableColumns;

            if(members.Length > 0)
            {
                foreach (var column in members)
                {
                    if (reader.GetOrdinal(column) != -1)
                    {
                        tableColumns.Add(new TableColumn()
                        {
                            Name = column,
                            Width = column.Length + 4
                        });
                    }
                }
            }
            else
            {
                for(int i=0; i<reader.FieldCount; i++)
                {
                    string colName = reader.GetName(i);

                    tableColumns.Add(new TableColumn()
                    {
                        Name = colName,
                        Width = colName.Length + 4
                    });
                }
            }

            

            return tableColumns;
        }

        /// <summary>
        /// Resizes columns and sets precalculated width
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columns"></param>
        /// <param name="colWidths"></param>
        private static void ResizeColumns(XSSFSheet sheet, IList<TableColumn> columns)
        {
            for (int i = 0; i < columns.Count; i++)
            {
                int width = (int)(columns[i].Width * 1.25f) * 256; // floating value can be adjusted at any time
                sheet.SetColumnWidth(i, Math.Min(width, 65279));
            }
        }

        private static void SetValueAndFormat(ICell cell, IDataFormat format, bool value)
        {
            cell.SetCellValue(value);
            SetCellStyle(cell, format, typeof(bool));
        }

        private static void SetValueAndFormat(ICell cell, IDataFormat format, double value)
        {
            cell.SetCellValue(value);
            SetCellStyle(cell, format, typeof(double));
        }

        private static void SetValueAndFormat(ICell cell, IDataFormat format, DateTime value)
        {
            cell.SetCellValue(value);
            SetCellStyle(cell, format, typeof(DateTime));
        }

        private static void SetValueAndFormat(ICell cell, IDataFormat format, string value)
        {
            cell.SetCellValue(value);
            SetCellStyle(cell, format, typeof(string));
        }

        private static void SetCellStyle(ICell cell, IDataFormat format, Type type)
        {
            ICellStyle cellStyle;

            if (_cellStyleCache.ContainsKey(type))
            {
                cellStyle = _cellStyleCache[type];
            }
            else
            {
                cellStyle = cell.Sheet.Workbook.CreateCellStyle();

                if (_cellFormatMapper.TryGetValue(type, out string formatString))
                {
                    short index = HSSFDataFormat.GetBuiltinFormat(formatString);

                    if (index != -1)
                    {
                        cellStyle.DataFormat = index;
                    }
                    else
                    {
                        cellStyle.DataFormat = format.GetFormat(formatString);
                    }
                }
                else
                {
                    cellStyle.DataFormat = format.GetFormat(DataFormats.TEXT);
                }

                _cellStyleCache.Add(type, cellStyle);
            }

            cell.CellStyle = cellStyle;
        }
    }

    internal class TableColumn
    {
        public string Name { get; set; }
        public int Width { get; set; }
    }
}
