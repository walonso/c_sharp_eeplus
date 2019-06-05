using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

namespace EntendiendoEPPLUS
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Unnamed_Click(object sender, EventArgs e)
        {
            string h = string.Empty;
            DataTable dt = GetTable();
            string rootapath = HttpContext.Current.Request.MapPath("~");
            string newFileName = string.Concat(rootapath, @"Temp\Plantilla - copia.xlsx");
            string realFileName = string.Concat(rootapath, @"Temp\Plantilla.xlsx");

            //Hacer pivot table (Tabla dinamica)
            CreateDataForPivotTable(dt, newFileName, "Data");
            CreatePivotTable(realFileName, "Data");
            

        }

        private void CreateChart(ExcelWorksheet ws, ExcelPivotTable pivotTable)
        {
            ExcelBarChart chart = ws.Drawings.AddChart("crtTiempos", eChartType.BarClustered, pivotTable) as ExcelBarChart;
            chart.SetPosition(1, 0, 4, 0);
            chart.SetSize(600, 400);

            chart.Title.Text = "Tiempos?";
            chart.Title.Font.Size = 18;
            chart.Title.Font.Bold = true;

            chart.GapWidth = 25;

            chart.DataLabel.ShowValue = true;

            chart.Legend.Remove();

            chart.XAxis.MajorTickMark = eAxisTickMark.None;
            chart.XAxis.MinorTickMark = eAxisTickMark.None;

            chart.YAxis.DisplayUnit = 1000; // K
            chart.YAxis.Deleted = true;

            ExcelBarChartSerie serie = chart.Series[0] as ExcelBarChartSerie;
            serie.Fill.Color = Color.FromArgb(91, 155, 213);

            chart.SetAxisGridlines(true, true, false);

            chart.SetCategoriesOrder(true);
        }
        public void CreatePivotTable(string curFileName, string sheetName)
        {
            using (ExcelPackage ep = new ExcelPackage(new FileInfo(curFileName)))
            {
                var wsData = ep.Workbook.Worksheets[sheetName];
                var tblData = wsData.Tables["tblData"];
                var dataCells = wsData.Cells[tblData.Address.Address];

                // workbook
                var wb = ep.Workbook;

                // new worksheet
                var ws = wb.Worksheets.Add("Tiempos por cliente");

                // default font
               // ws.Cells.Style.Font.SetFromFont(font);

                // cells borders
                ws.View.ShowGridLines = false;

                // tab color
                ws.TabColor = Color.FromArgb(91, 155, 213);

                // columns width
                ws.Column(1).Width = 2.5;

                ExcelPivotTable pivotTable = ws.PivotTables.Add(ws.Cells["B4"], dataCells, "pvtTiemposPorCliente");

                // headers
                pivotTable.ShowHeaders = true;
                pivotTable.RowHeaderCaption = "Cliente";

                // grand total
                pivotTable.ColumGrandTotals = true;
                pivotTable.GrandTotalCaption = "Total";

                // data fields are placed in columns
                pivotTable.DataOnRows = false;

                // style
                pivotTable.TableStyle = OfficeOpenXml.Table.TableStyles.Medium9;

                ExcelPivotTableField territoryGroupPageField = pivotTable.PageFields.Add(pivotTable.Fields["Empresa"]);
                territoryGroupPageField.Sort = eSortType.Ascending;

                ExcelPivotTableField salesPersonRowField = pivotTable.RowFields.Add(pivotTable.Fields["Cliente"]);

                ExcelPivotTableDataField revenueDataField = pivotTable.DataFields.Add(pivotTable.Fields["Tiempo"]);
                revenueDataField.Function = DataFieldFunctions.Sum;
                revenueDataField.Format = string.Format("{0};{1}", "#,##0_)", "(#,##0)");
                revenueDataField.Name = "Tiempo";

                pivotTable.SortOnDataField(salesPersonRowField, revenueDataField, true);
                pivotTable.Top10(salesPersonRowField, revenueDataField, 5, false,false);

                //Crear el chart table
                CreateChart(ws, pivotTable);



                #region TemporalPruebas
                string rootapath = HttpContext.Current.Request.MapPath("~");
                string newFileName = string.Concat(rootapath, @"Temp\Plantilla_Ralenti_NO_WT.xlsx");

                FileInfo excelFile = new FileInfo(newFileName);
                if (excelFile.Exists)
                    excelFile.Delete();
                ep.SaveAs(excelFile);
                #endregion
                //AdventureWorks8_PivotChart_SalesBySalesperson(ws, pivotTable);
            }
        }

        public void CreateDataForPivotTable(DataTable dt, string curFileName, string sheetName)
        {
            using (ExcelPackage ep = new ExcelPackage(new FileInfo(curFileName)))
            {

                // rows and columns indices
                int startRowIndex = 2;
                int empresaGroupIndex = 2;
                int clienteNameIndex = 3;
                int identificadorIndex = 4;
                int DateIndex = 5;
                int horaIndex = 6;
                int tiempoMaxIndex = 7;

                // workbook
                var wb = ep.Workbook;

                // new worksheet
                var ws = wb.Worksheets.Add(sheetName);

                // default font
                //ws.Cells.Style.Font.SetFromFont(new Font("Tahoma", 10););

                // cells borders
                ws.View.ShowGridLines = false;

                // load data
                using (ExcelRangeBase range = ws.Cells[startRowIndex, empresaGroupIndex].LoadFromDataTable(dt, true, TableStyles.Medium2))
                {
                    // border style
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    // border color
                    range.Style.Border.Top.Color.SetColor(Color.LightGray);
                    range.Style.Border.Bottom.Color.SetColor(Color.LightGray);
                    range.Style.Border.Left.Color.SetColor(Color.LightGray);
                    range.Style.Border.Right.Color.SetColor(Color.LightGray);
                }

                // data table
                // LoadFromCollection adds Excel Table
                ExcelTable tblData = ws.Tables[ws.Tables.Count - 1];
                tblData.Name = "tblData";

                // headers
                ws.Cells[startRowIndex, empresaGroupIndex].Value = "Empresa";
                ws.Cells[startRowIndex, clienteNameIndex].Value = "Cliente";
                ws.Cells[startRowIndex, identificadorIndex].Value = "Identificador";
                ws.Cells[startRowIndex, DateIndex].Value = "Fecha Date";
                ws.Cells[startRowIndex, horaIndex].Value = "Hora";
                ws.Cells[startRowIndex, tiempoMaxIndex].Value = "Tiempo";

                // headers style
                using (var cells = ws.Cells[startRowIndex, empresaGroupIndex, startRowIndex, tiempoMaxIndex])
                {
                    cells.Style.Font.Bold = true;
                    cells.Style.Font.Size = 11;
                    cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                int fromRowIndex = startRowIndex + 1;
                int toRowIndex = startRowIndex + dt.Rows.Count;

                // cells format & horizontal alignment
                ws.Cells[fromRowIndex, empresaGroupIndex, toRowIndex, tiempoMaxIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                using (var cells = ws.Cells[fromRowIndex, DateIndex, toRowIndex, DateIndex])
                {
                    cells.Style.Numberformat.Format = "dd/mm/yyyy";
                    cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }

                //using (var cells = ws.Cells[fromRowIndex, orderQtyIndex, toRowIndex, orderQtyIndex])
                //{
                //    cells.Style.Numberformat.Format = numberIntFormat;
                //    cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //}

                //using (var cells = ws.Cells[fromRowIndex, unitPriceDiscountIndex, toRowIndex, lineTotalIndex])
                //{
                //    cells.Style.Numberformat.Format = numberFormat;
                //    cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                //}

                // tab color
                ws.TabColor = Color.FromArgb(128, 0, 0);

                // columns width
                for (int columnIndex = 1; columnIndex < empresaGroupIndex; columnIndex++)
                    ws.Column(columnIndex).Width = 2.5;
                for (int columnIndex = empresaGroupIndex; columnIndex <= tiempoMaxIndex; columnIndex++)
                {
                    ws.Column(columnIndex).AutoFit();
                    ws.Column(columnIndex).Width += 2.5;
                }

                #region TemporalPruebas
                string rootapath = HttpContext.Current.Request.MapPath("~");
                string newFileName = string.Concat(rootapath, @"Temp\Plantilla.xlsx");

                FileInfo excelFile = new FileInfo(newFileName);
                if (excelFile.Exists)
                    excelFile.Delete();
                ep.SaveAs(excelFile);
                #endregion

            }
        }



        public DataTable GetTable()
        {

            DataTable table = new DataTable();
            table.Columns.Add("Carro", typeof(string));
            table.Columns.Add("Conductor", typeof(string));
            table.Columns.Add("Placa", typeof(string));
            table.Columns.Add("Fecha", typeof(DateTime));
            table.Columns.Add("Total", typeof(TimeSpan));
            table.Columns.Add("Tiempo", typeof(TimeSpan));


            table.Rows.Add("LIV", "JOS JER", "2", DateTime.Now, "00:05:06", "00:05:06");
            table.Rows.Add("LIV TEB", " ", "1", DateTime.Now, "00:05:06", "00:05:06");
            table.Rows.Add("TEB", "JESUS R", "3", DateTime.Now, "00:06:06", "00:06:06");
            table.Rows.Add("LIVTEB", "HOOV SEGURO", "3", DateTime.Now, "00:06:06", "00:06:06");
            table.Rows.Add("LIV W", "NICOL ESC", "3", DateTime.Now, "00:07:06", "00:07:06");
            table.Rows.Add("LIV A TEB", "HAR OSP", "3", DateTime.Now, "00:07:06", "00:07:06");
            table.Rows.Add("LIV B TEB", "DAL MUR", "3", DateTime.Now, "00:08:06", "00:08:06");
            table.Rows.Add("LIV C TEB", "JOS CAN", "3", DateTime.Now, "00:09:06", "00:09:06");
            table.Rows.Add("LIV D TEB", "AL VAS", "3", DateTime.Now, "00:09:06", "00:09:06");


            return table;
        }
    }
}