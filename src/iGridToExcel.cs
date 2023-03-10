using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Media;
using TenTec.Windows.iGridLib;

namespace iGridToExcel
{
    public static class iGridToExcel
    {
        public static (string, string) ExportSingleGrid(string workbook_title, string worksheet_label, string headerText, iGrid ig) // returns filename,error message
        {
            return ExportMultiGrids(workbook_title, worksheet_label, headerText, new List<string>(), new List<iGrid>() { ig });
        }

        public static (string, string) ExportMultiGrids(string workbook_title, string worksheet_label, string headerText, List<string> gridTitles, List<iGrid> igs) // returns filename,error message
        {
            try
            {
                ExcelPackage pck = null;
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                pck = new ExcelPackage();
                ExcelWorkbook wb = pck.Workbook;
                if (!string.IsNullOrEmpty(workbook_title)) wb.Properties.Title = workbook_title;

                if (string.IsNullOrEmpty(worksheet_label)) worksheet_label = "Sheet1";
                ExcelWorksheet ws = wb.Worksheets.Add(worksheet_label);
                int curRow = 1;
                if (!string.IsNullOrEmpty(headerText))
                {
                    foreach (string ln in headerText.Split('\n'))
                    {
                        ws.Cells[curRow, 1].Value = ln;
                        curRow += 1;
                    }
                    curRow += 1;
                }
                if (igs == null || igs.Count == 0) goto Done;

                string err = "";
                for (int i = 0; i < igs.Count; i++)
                {
                    string hText = "";
                    if (i < gridTitles.Count) hText = gridTitles[i];
                    (curRow, err) = WriteGridToExcel(curRow, ws, hText, igs[i]);
                    if (!string.IsNullOrEmpty(err)) Debugger.Break();
                    curRow++;
                }


                Done:
                string excelTmpFile = Path.GetTempFileName().ToLower().Replace(".tmp", ".xlsx");
                pck.SaveAs(new FileInfo(excelTmpFile));
                pck.Dispose();
                return (excelTmpFile, "");
            }

            catch (Exception ex)
            {
                Debugger.Break();
                return ("", ex.Message);
            }
        }

        private static (int, string) WriteGridToExcel(int curRow, ExcelWorksheet ws, string headerText, iGrid ig)
        {
            try
            {
                System.Drawing.Color col = System.Drawing.Color.Empty;

                if (!string.IsNullOrEmpty(headerText))
                {
                    ws.Cells[curRow, 1].Value = headerText;
                    ws.Cells[curRow, 1].Style.Font.Size = 16;
                    ws.Cells[curRow, 1].Style.Font.Bold = true;
                    curRow += 1;
                }

                // check for multiple headers
                for (int i = ig.Header.Rows.Count - 1; i >= 0; i--)
                {
                    int jCol = 0;
                    for (int j = 0; j < ig.Cols.Count; j++)
                    {
                        if (!ig.Cols[j].Visible) continue;
                        jCol += 1;
                        iGColHdr colHdr = ig.Header.Cells[i, j];
                        if (colHdr.Value == null)
                            continue;

                        double d;
                        bool b = Double.TryParse(colHdr.Value.ToString(), out d);
                        if (b)
                            ws.Cells[curRow, jCol].Value = colHdr.Value;
                        else
                            ws.Cells[curRow, jCol].Value = colHdr.Value.ToString();

                        if (colHdr.BackColor != null && colHdr.BackColor != System.Drawing.Color.Empty
                            && colHdr.BackColor != System.Drawing.Color.White && colHdr.BackColor != System.Drawing.Color.Transparent)
                            for (int k = 0; k < colHdr.SpanCols; k++)
                            {
                                ws.Cells[curRow, jCol + k].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                col = colHdr.BackColor;
                                if (col.A > 0 && col.A < 255)
                                    col = MixColoursWeight(col, col.A, System.Drawing.Color.White, 255 - col.A);
                                ws.Cells[curRow, jCol + k].Style.Fill.BackgroundColor.SetColor(col);
                            }
                        ws.Cells[curRow, jCol].Style.HorizontalAlignment = ExcelAlignment(colHdr.TextAlign);
                        if (colHdr.Font != null)
                        {
                            ws.Cells[curRow, jCol].Style.Font.Size = colHdr.Font.Size;
                            if (colHdr.Font.Bold) ws.Cells[curRow, jCol].Style.Font.Bold = true;
                            if (colHdr.Font.Italic) ws.Cells[curRow, jCol].Style.Font.Italic = true;
                            if (colHdr.Font.Underline) ws.Cells[curRow, jCol].Style.Font.UnderLine = true;
                        }
                        if (colHdr.ForeColor != System.Drawing.Color.Black) ws.Cells[curRow, jCol].Style.Font.Color.SetColor(colHdr.ForeColor);

                        ws.Cells[curRow, jCol, curRow, jCol + colHdr.SpanCols - 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    }
                    curRow += 1;
                }

                // write data
                col = System.Drawing.Color.Empty;
                for (int i = 0; i < ig.Rows.Count; i++)
                {
                    int jCol = 0;
                    for (int j = 0; j < ig.Cols.Count; j++)
                    {
                        if (!ig.Cols[j].Visible) continue;
                        iGCell cell = ig.Rows[i].Cells[j];
                        jCol += 1;
                        if (cell.Value == null) continue;

                        double d;
                        bool b = Double.TryParse(cell.Value.ToString(), out d);
                        if (b)
                            ws.Cells[curRow, jCol].Value = cell.Value;
                        else
                            ws.Cells[curRow, jCol].Value = cell.Value.ToString();

                        col = System.Drawing.Color.Empty;
                        if (ig.Cols[j].CellStyle.BackColor != col)
                            col = ig.Cols[j].CellStyle.BackColor;
                        if (cell.BackColor != System.Drawing.Color.Empty && cell.BackColor != col)
                            col = cell.BackColor;

                        if (col != System.Drawing.Color.Empty)
                        {
                            if (col.A > 0 && col.A < 255)
                                col = MixColoursWeight(col, col.A, System.Drawing.Color.White, 255 - col.A);
                            ws.Cells[curRow, jCol].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[curRow, jCol].Style.Fill.BackgroundColor.SetColor(col);
                        }

                        ws.Cells[curRow, jCol].Style.HorizontalAlignment = ExcelAlignment(ig.Cols[j].CellStyle.TextAlign);
                        if (cell.Font != null)
                        {
                            ws.Cells[curRow, jCol].Style.Font.Size = cell.Font.Size;
                            if (cell.Font.Bold) ws.Cells[curRow, jCol].Style.Font.Bold = true;
                            if (cell.Font.Italic) ws.Cells[curRow, jCol].Style.Font.Italic = true;
                            if (cell.Font.Underline) ws.Cells[curRow, jCol].Style.Font.UnderLine = true;
                        }


                        col = System.Drawing.Color.Empty;
                        if (ig.Cols[j].CellStyle.ForeColor != col)
                            col = ig.Cols[j].CellStyle.ForeColor;
                        if (cell.ForeColor != System.Drawing.Color.Empty && cell.ForeColor != col)
                            col = cell.ForeColor;
                        if (col != System.Drawing.Color.Empty) ws.Cells[curRow, jCol].Style.Font.Color.SetColor(col);

                        ws.Cells[curRow, jCol].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                    }
                    curRow += 1;
                }


                return (curRow, "");
            }
            catch (Exception ex)
            {
                Debugger.Break();
                return (curRow, ex.Message);
            }
        }

        private static OfficeOpenXml.Style.ExcelHorizontalAlignment ExcelAlignment(iGContentAlignment aln)
        {
            switch (aln)
            {
                case iGContentAlignment.TopLeft:
                case iGContentAlignment.MiddleLeft:
                case iGContentAlignment.BottomLeft:
                    return OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                case iGContentAlignment.TopRight:
                case iGContentAlignment.MiddleRight:
                case iGContentAlignment.BottomRight:
                    return OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                case iGContentAlignment.TopCenter:
                case iGContentAlignment.MiddleCenter:
                case iGContentAlignment.BottomCenter:
                    return OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                default:
                    return OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            }
        }
        private static System.Drawing.Color MixColoursWeight(System.Drawing.Color col1, double w1, System.Drawing.Color col2, double w2)
        {
            if (w1 == 0 & w2 == 0)
            {
                w1 = 1;
                w2 = 1;
            }
            int a, r, g, b;
            a = System.Convert.ToInt32((System.Convert.ToDouble(col1.A) * w1 + System.Convert.ToDouble(col2.A) * w2) / (w1 + w2));
            r = System.Convert.ToInt32((System.Convert.ToDouble(col1.R) * w1 + System.Convert.ToDouble(col2.R) * w2) / (w1 + w2));
            g = System.Convert.ToInt32((System.Convert.ToDouble(col1.G) * w1 + System.Convert.ToDouble(col2.G) * w2) / (w1 + w2));
            b = System.Convert.ToInt32((System.Convert.ToDouble(col1.B) * w1 + System.Convert.ToDouble(col2.B) * w2) / (w1 + w2));
            if (a > 255)
            {
                a = 255; if (a < 0)
                    a = 0;
            }
            if (r > 255)
            {
                r = 255; if (r < 0)
                    r = 0;
            }
            if (g > 255)
            {
                g = 255; if (g < 0)
                    g = 0;
            }
            if (b > 255)
            {
                b = 255; if (b < 0)
                    b = 0;
            }

            return System.Drawing.Color.FromArgb(a, r, g, b);
        }
    }


}

