using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ConsoleApp4.Excel
{
    public class PivotHelper
    {
        private static readonly DataConsolidateFunction DefaultDCF = DataConsolidateFunction.COUNT;

        /// <summary>
        /// Creates a new sheet and adds a dynamic table based on source table. 
        /// Contains no columns by default.
        /// </summary>
        /// <param name="table"></param>
        /// <param name="settings"></param>
        /// <returns><see cref="XSSFPivotTable"/></returns>
        public static XSSFPivotTable CreatePivotTable(XSSFTable table, PivotSettings settings)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            // Get Table Range
            var startReference = new CellReference(table.StartRowIndex, table.StartColIndex);
            var endReference = new CellReference(table.EndRowIndex, table.EndColIndex);
            var range = new AreaReference(startReference, endReference);

            // Create new sheet in the current workbook
            XSSFSheet sourceSheet = table.GetXSSFSheet();
            IWorkbook workbook = sourceSheet.Workbook;
            XSSFSheet pivotSheet = (XSSFSheet)workbook.CreateSheet(settings.SheetName);

            // Create Pivot Table
            var pivotTablePosition = new CellReference(0, 0);

            XSSFPivotTable pivotTable = pivotSheet.CreatePivotTable(range, pivotTablePosition, sourceSheet);
            CT_PivotTableDefinition pivotTableDef = pivotTable.GetCTPivotTableDefinition();
            pivotTableDef.name = settings.TableName;

            AddRowLabels(table, pivotTable, settings.RowLabels);
            AddColumnLabels(table, pivotTable, settings.ColumnLabels);
            AddFilterLabels(table, pivotTable, settings.FilterLabels);
            SetPivotTableStyle(pivotTable, settings.TableStyle);
            CreateFreezePane(pivotTable);

            return pivotTable;
        }

        /// <summary>
        /// Collapses row and column fields in a pivot table.
        /// Constructs the pivot cache by itterating over the column values.
        /// </summary>
        /// <param name="table"></param>
        /// <param name="pivotTable"></param>
        /// <param name="columnIndex"></param>
        private static void CollapseFields(XSSFTable table, XSSFPivotTable pivotTable, int columnIndex)
        {
            // Get all unique values in a column
            var colAValues = new HashSet<string>();

            /* ignore header row */
            for (int r = table.StartRowIndex + 1; r <= table.EndRowIndex; r++)
            {
                IRow row = table.GetXSSFSheet().GetRow(r);

                if (row != null)
                {
                    ICell cell = row.GetCell(columnIndex);

                    if (cell != null)
                    {
                        colAValues.Add(cell.ToString());
                    }
                }
            }

            /* EXTREMELY LOW LEWEL. DO NOT TOUCH THIS CODE */

            var items = pivotTable.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(columnIndex).items.item;
            int i = 0;

            CT_Item ct_item;

            foreach (string col in colAValues)
            {
                ct_item = items[i];
                ct_item.t = ST_ItemType.data;
                ct_item.x = (uint)i++;

                var sharedItem = new CT_String { v = col };

                pivotTable.GetPivotCacheDefinition()
                    .GetCTPivotCacheDefinition().cacheFields
                    .cacheField[columnIndex].sharedItems.Items
                    .Add(sharedItem);
                ct_item.sd = false;
            }

            while (i < colAValues.Count)
            {
                ct_item = items[i++];
                ct_item.sd = false;
            }

            /* END OF LOW LEVEL CODE */
        }

        /// <summary>
        /// Adds column labels to <see cref="XSSFPivotTable"/>
        /// </summary>
        /// <param name="table"></param>
        /// <param name="pivotTable"></param>
        /// <param name="labels"></param>
        private static void AddColumnLabels(XSSFTable table, XSSFPivotTable pivotTable, List<PivotSettings.PivotColumnLabel> labels)
        {
            foreach (var item in labels)
            {
                int ix = table.FindColumnIndex(item.Name);
                
                if (ix != -1)
                    pivotTable.AddColumnLabel(item.DataConsolidateFunction ?? DefaultDCF, ix, item.Name);
            }
        }

        /// <summary>
        /// Adds filter labels to <see cref="XSSFTable"/>
        /// </summary>
        /// <param name="table"></param>
        /// <param name="pivotTable"></param>
        /// <param name="columnNames"></param>
        private static void AddFilterLabels(XSSFTable table, XSSFPivotTable pivotTable, string[] columnNames)
        {
            foreach (string item in columnNames)
            {
                int ix = table.FindColumnIndex(item);

                if (ix != -1)
                    pivotTable.AddReportFilter(ix);
            }
        }

        /// <summary>
        /// Adds row labels to <see cref="XSSFTable"/>
        /// </summary>
        /// <param name="table"></param>
        /// <param name="pivotTable"></param>
        /// <param name="labels"></param>
        private static void AddRowLabels(XSSFTable table, XSSFPivotTable pivotTable, List<PivotSettings.PivotRowLabel> labels)
        {
            CT_PivotTableDefinition pivotTableDef = pivotTable.GetCTPivotTableDefinition();
            CT_ColFields colFields = null;

            foreach (var label in labels)
            {
                int index = table.FindColumnIndex(label.Name);

                if (index == -1)
                    continue;

                pivotTable.AddRowLabel(index);
                CT_PivotField pivotField = pivotTable.GetCTPivotTableDefinition().pivotFields.GetPivotFieldArray(index);
                pivotField.sortType = label.SortType;
                pivotField.showAll = false;

                /* LOW LEVEL FLIPPING OF ROW LABEL  AXIS */

                CT_RowFields rowFieldsObj = pivotTableDef.rowFields;

                if (label.Direction == ST_Axis.axisCol)
                {
                    colFields ??= pivotTableDef.AddNewColFields();
                    List<CT_Field> rowFields = rowFieldsObj.field;
                    rowFields.RemoveAt(rowFields.FindIndex(f => f.x == index));
                    rowFieldsObj.count--;
                    colFields.AddNewField().x = index;
                    pivotTableDef.pivotFields
                        .GetPivotFieldArray(index).axis = ST_Axis.axisCol;
                    pivotTableDef.colFields.count += 1;
                }

                CollapseFields(table, pivotTable, index);
            }
        }

        /// <summary>
        /// Styles the pivot table.
        /// </summary>
        /// <param name="table"></param>
        /// <param name="tableStyle"></param>
        private static void SetPivotTableStyle(XSSFPivotTable table, string tableStyle)
        {
            CT_PivotTableDefinition ptDef = table.GetCTPivotTableDefinition();

            if (tableStyle != null)
            {
                ptDef.pivotTableStyleInfo.name = tableStyle;
            }

            ptDef.compact = true;
            ptDef.compactData = true;
            ptDef.outline = true;
            ptDef.outlineData = true;

            if (ptDef.pivotFields != null)
            {
                foreach (var item in ptDef.pivotFields.pivotField)
                {
                    item.showAll = false;
                    item.topAutoShow = false;
                    item.compact = true;
                }
            }
        }

        private static void CreateFreezePane(XSSFPivotTable pivotTable)
        {
            CT_PivotTableDefinition ct_table = pivotTable.GetCTPivotTableDefinition();
            int firstColIndex = (int)ct_table.location.firstDataCol;
            int firstRowIndex = (int)ct_table.location.firstDataRow;
            int colCount = (int) ct_table.colFields.count;
            pivotTable.GetParentSheet().CreateFreezePane(firstColIndex, firstRowIndex + colCount);
        }
    }
}
