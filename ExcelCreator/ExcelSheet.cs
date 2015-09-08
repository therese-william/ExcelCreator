using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCreator
{
    /// <summary>
    /// Define Excel file sheet
    /// </summary>
    public class ExcelSheet
    {
        string sheetName;
        /// <summary>
        /// Set or Get the name of the sheet as it will appear in the excel file
        /// </summary>
        public string SheetName
        {
            get { return sheetName; }
            set { sheetName = value; }
        }

        List<ExcelColumn> columns;
        /// <summary>
        /// Set or Get the columns of this sheet
        /// </summary>
        public List<ExcelColumn> Columns
        {
            get { return columns; }
            set { columns = value; }
        }

        List<List<string>> rows;

        /// <summary>
        /// Get the rows included in this sheet
        /// </summary>
        public List<List<string>> Rows
        {
            get { return rows; }
            //set { rows = value; }
        }

        /// <summary>
        /// Add column to sheet by providing a column of type ExcelColumn
        /// </summary>
        /// <param name="col">Column to be added to the sheet</param>
        public void AddColumn(ExcelColumn col){
            if(columns == null){
                columns = new List<ExcelColumn>();
            }
            columns.Add(col);
        }
        /// <summary>
        /// Add column to sheet by providing the column details.
        /// </summary>
        /// <param name="ColumnName">Column Name</param>
        /// <param name="ColumnType">Column Type is Text, DropDown or MultiSelect</param>
        /// <param name="ColumnOptions">If column is DropDown or MultiSelect then Column Options should be supplied as a list of strings</param>
        public void AddColumn(string ColumnName, ExcelColumn.ColumnTypes ColumnType, List<string> ColumnOptions, int ColumnWidth)
        {
            ExcelColumn col = new ExcelColumn();
            col.ColumnName = ColumnName;
            col.ColumnOptions = ColumnOptions;
            col.ColumnType = ColumnType;
            col.ColumnWidth = ColumnWidth;
            if (columns == null)
            {
                columns = new List<ExcelColumn>();
            }
            columns.Add(col);
        }

        /// <summary>
        /// Add row to sheet, given that the number of submitted values should be the same as added columns, so for empty columns you can just add empty string
        /// </summary>
        /// <param name="Values">List of values as strings</param>
        public void AddRow(List<string> Values)
        {
            if (rows == null)
            {
                rows = new List<List<string>>();
            }

            if (Values != null && columns != null && Values.Count == columns.Count)
            {
                rows.Add(Values);
            }
            else
            {
                throw new InvalidSheetException("Row values don't match columns count, please insert empty string in balnk columns.");
            }
        }
    }
}
