using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCreator
{
    /// <summary>
    /// Define column for excel sheet
    /// </summary>
    public class ExcelColumn
    {
        public enum ColumnTypes { Text, DropDown, MultiSelect };
        string columnName;
        /// <summary>
        /// Name of the column
        /// </summary>
        public string ColumnName
        {
            get { return columnName; }
            set { columnName = value; }
        }
        ColumnTypes columnType;
        /// <summary>
        /// Column type is Text, DropDown or MultiSelect
        /// </summary>
        public ColumnTypes ColumnType
        {
            get { return columnType; }
            set { columnType = value; }
        }
        /// <summary>
        /// If column is DropDown or MultiSelect then Column Options should be supplied as a list of strings
        /// </summary>
        List<string> columnOptions;

        public List<string> ColumnOptions
        {
            get { return columnOptions; }
            set { columnOptions = value; }
        }

        int columnWidth;

        public int ColumnWidth
        {
            get { return columnWidth; }
            set { columnWidth = value; }
        }

    }
}
