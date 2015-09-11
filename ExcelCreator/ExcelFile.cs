using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Configuration;
using VBIDE = Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelCreator
{
    /// <summary>
    /// Excel File
    /// </summary>
    public class ExcelFile
    {
        const string _MultiSelectProcedure = 
            "Private Sub Worksheet_Change(ByVal Target As Range)\n"+
            "' Developed by Contextures Inc.\n"+
            "' www.contextures.com\n"+
            "Dim rngDV As Range\n"+
            "Dim oldVal As String\n"+
            "Dim newVal As String\n"+
            "If Target.Count > 1 Then GoTo exitHandler\n"+
            "On Error Resume Next\n"+
            "Set rngDV = Cells.SpecialCells(xlCellTypeAllValidation)\n"+
            "On Error GoTo exitHandler\n"+
            "If rngDV Is Nothing Then GoTo exitHandler\n"+
            "If Intersect(Target, rngDV) Is Nothing Then\n"+
            "   'do nothing\n"+
            "Else\n"+
            "  Application.EnableEvents = False\n"+
            "  newVal = Target.Value\n"+
            "  Application.Undo\n"+
            "  oldVal = Target.Value\n"+
            "  Target.Value = newVal\n"+
            "  If UBound(Filter(Array({0}), Target.Column)) <> -1 Then\n" +
            "      If (InStr(oldVal, newVal & \",\") > 0 Or InStr(oldVal, \",\" & newVal) > 0 Or newVal = oldVal) And newVal <> \"\" Then\n" +
            "        Target.Value = oldVal\n"+
            "      Else\n"+
            "        If oldVal = \"\" Then\n"+
            "          'do nothing\n"+
            "          Else\n"+
            "          If newVal = \"\" Then\n"+
            "            'do nothing\n"+
            "          Else\n"+
            "          Target.Value = oldVal _\n"+
            "            & \",\" & newVal\n"+
            "    '      NOTE: you can use a line break,\n"+
            "    '      instead of a comma\n"+
            "    '      Target.Value = oldVal _\n"+
            "    '        & Chr(10) & newVal\n"+
            "          End If\n"+
            "        End If\n"+
            "      End If\n"+
            "    End If\n"+
            "End If\n"+
            "exitHandler:\n"+
            "  Application.EnableEvents = True\n"+
            "End Sub";

        string fileName;
        /// <summary>
        /// File name as it will be saved
        /// </summary>
        public string FileName
        {
            get { return fileName; }
            set { fileName = value; }
        }

        List<ExcelSheet> sheets;
        /// <summary>
        /// Sheets included in excel file
        /// </summary>
        public List<ExcelSheet> Sheets
        {
            get { return sheets; }
            set { sheets = value; }
        }

        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;

        /// <summary>
        /// Sets border of cells from beginCell to endCell to a single-line width border all around.
        /// </summary>
        /// <param name="ws">The worksheet to modify cells in</param>
        /// <param name="beginCell">The cell to begin the border change at, e.g. "A1"</param>
        /// <param name="endCell">The cell to end the border change at, e.g. "C2"</param>
        private void encloseCellWithBorder(Excel.Worksheet ws, string beginCell, string endCell)
        {
            Excel.Range oResizeRange = ws.get_Range(beginCell, endCell);
            oResizeRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
        }

        /// <summary>
        /// Sets color of cells from beginCell to endCell to the specified color.
        /// </summary>
        /// <param name="ws">The worksheet to modify cells in</param>
        /// <param name="beginCell">The cell to begin the color change at, e.g. "A1"</param>
        /// <param name="endCell">The cell to end the color change at, e.g. "C2"</param>
        /// <param name="color">The System.Drawing.Color color to give the cell interior</param>
        private void setInterior(Excel.Worksheet ws, string beginCell, string endCell, System.Drawing.Color color)
        {
            Excel.Range oResizeRange = ws.get_Range(beginCell, endCell);
            oResizeRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
            oResizeRange.Font.Bold = true;
        }
        /// <summary>
        /// Save Excel to file
        /// </summary>
        /// <returns>Full path of created excel file</returns>
        public string SaveToExcelFile()
        {
            string fullPath = "";
            if (sheets.Count > 0)
            {
                int excelActiveRow;
                //initialize the Excel application and make it invisible to the user.
                xlApp = new Excel.Application();

                //This is a must for lengthy spreadsheets - the Excel COM library does not like it if a user clicks around in the spreadsheet while it is being created.
                xlApp.UserControl = false;
                xlApp.Visible = false;

                //Create the Excel workbook and worksheet - and give the worksheet a name.
                xlWorkBook = (Excel.Workbook)(xlApp.Workbooks.Add(Missing.Value));

                int SheetNumber = 1;


                foreach (ExcelSheet sheet in sheets)
                {
                    List<int> multiSelectCols = new List<int>();

                    Excel.Worksheet xlWorkSheet;
                    if (SheetNumber == 1)
                    {
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
                    }
                    else
                    {
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.Add();
                        ((Microsoft.Office.Interop.Excel._Worksheet)xlWorkSheet).Activate();
                    }

                    xlWorkSheet.Name = sheet.SheetName;

                    if (sheet.Columns != null && sheet.Columns.Count > 0)
                    {

                        //the first row of the excel spreadsheet is row 1. We maintain the current active row of the spreadsheet as a variable for easy tracking.
                        excelActiveRow = 1;

                        int ColumnNumber = 1;
                        string ColumnChar="";
                        //Add defined sheet columns to first row
                        foreach (ExcelColumn c in sheet.Columns)
                        {
                            ColumnChar = Char.ConvertFromUtf32(ColumnNumber + 64);
                            //to insert a value into an Excel cell, use the .Cells[Row, column] property of the Excel.WorkSheet class.
                            xlWorkSheet.Cells[excelActiveRow, ColumnNumber] = c.ColumnName;
                            if ((c.ColumnType == ExcelColumn.ColumnTypes.DropDown || c.ColumnType == ExcelColumn.ColumnTypes.MultiSelect) && c.ColumnOptions != null)
                            {
                                //Add validation to all the column cells
                                xlWorkSheet.get_Range(ColumnChar + "2", ColumnChar + "1048576").Validation.Delete();
                                string flatList = string.Join(",", c.ColumnOptions.ToArray());
                                xlWorkSheet.get_Range(ColumnChar + "2", ColumnChar + "1048576").Validation
                                    .Add(XlDVType.xlValidateList, Type.Missing,XlFormatConditionOperator.xlBetween,flatList,Type.Missing);

                                if (c.ColumnType == ExcelColumn.ColumnTypes.MultiSelect)
                                {
                                    multiSelectCols.Add(ColumnNumber);
                                }
                            }

                            if (c.ColumnWidth > 0)
                            {
                                xlWorkSheet.Columns[ColumnNumber].ColumnWidth = c.ColumnWidth;
                            }
                            if (c.WrapText)
                            {
                                xlWorkSheet.Columns[ColumnNumber].WrapText = true;
                            }
                            ColumnNumber++;
                        }

                        setInterior(xlWorkSheet, "A1", ColumnChar + "1", Color.Gray);
                        encloseCellWithBorder(xlWorkSheet, "A1", ColumnChar + "1048576");

                        //Add defined rows
                        if (sheet.Rows != null && sheet.Rows.Count > 0)
                        {
                            int RowNumber = 2;
                            foreach (List<string> row in sheet.Rows)
                            {
                                if (row != null && row.Count == sheet.Columns.Count)
                                {
                                    ColumnNumber = 1;
                                    foreach (string colval in row)
                                    {
                                        xlWorkSheet.Cells[RowNumber, ColumnNumber] = colval;
                                        ColumnNumber++;
                                    }
                                }

                                RowNumber++;
                            }
                        }
                    }

                    if (sheet.Columns.Exists(c => c.ColumnType == ExcelColumn.ColumnTypes.MultiSelect))
                    {
                        string multiselectcols = string.Join(",", multiSelectCols);
                        var excelsheet = xlWorkBook.VBProject.VBComponents.OfType<VBIDE.VBComponent>().Where(c => c.Name == "Sheet" + SheetNumber.ToString()).ToList().FirstOrDefault();
                        excelsheet.CodeModule.AddFromString(String.Format(_MultiSelectProcedure,multiselectcols));
                    }

                    SheetNumber++;

                }

                //save the workbook to "C:\workbook.xlsm";
                fullPath = Properties.Settings.Default.FileDirectory + fileName + ".xlsm";
                if (File.Exists(fullPath))
                {
                    File.Delete(fullPath);
                }

                try
                {//save the workbook.
                    xlWorkBook.SaveAs(fullPath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
                    null, null, false, false, Excel.XlSaveAsAccessMode.xlNoChange, false, false, null, null, null);
                }
                catch (Exception ex)
                {
                    //release all memory - stop EXCEL.exe from hanging around.
                    //if (xlMod != null) { Marshal.ReleaseComObject(xlMod); }
                    if (xlWorkBook != null) { Marshal.ReleaseComObject(xlWorkBook); }
                    //if (xlWorkSheet != null) { Marshal.ReleaseComObject(xlWorkSheet); }
                    if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }
                    //xlMod = null;
                    xlWorkBook = null;
                    //xlWorkSheet = null;
                    xlApp = null;
                    GC.Collect();

                    Environment.Exit(0);
                }


                xlApp.Quit();

                //release all memory - stop EXCEL.exe from hanging around.
                //if (xlMod != null) { Marshal.ReleaseComObject(xlMod); }
                if (xlWorkBook != null) { Marshal.ReleaseComObject(xlWorkBook); }
                //if (xlWorkSheet != null) { Marshal.ReleaseComObject(xlWorkSheet); }
                if (xlApp != null) { Marshal.ReleaseComObject(xlApp); }
                //xlMod = null;
                xlWorkBook = null;
                //xlWorkSheet = null;
                xlApp = null;
                GC.Collect();
            }
            else
            {
                throw new EmptyExcelException("No sheets added to the file!");
            }
            return fullPath;
        }
    }
}
