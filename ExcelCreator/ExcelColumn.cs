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
        /// <summary>
        /// Column Types used to specify if this column is just text or drop down list or multi select
        /// </summary>
        public enum ColumnTypes { 
            /// <summary>
            /// Text column, user can add any text to it
            /// </summary>
            Text, 
            /// <summary>
            /// Drop down column user should choose only one value from the set of values defined in ColumnOptions field
            /// </summary>
            DropDown, 
            /// <summary>
            /// Multi select column user should choose many values from the set of values defined in ColumnOptions field
            /// </summary>
            MultiSelect };
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
        List<string> columnOptions;

        /// <summary>
        /// If column is DropDown or MultiSelect then Column Options should be supplied as a list of strings
        /// </summary>
        public List<string> ColumnOptions
        {
            get { return columnOptions; }
            set { columnOptions = value; }
        }
        int columnWidth;

        /// <summary>
        /// Width of the column
        /// </summary>
        public int ColumnWidth
        {
            get { return columnWidth; }
            set { columnWidth = value; }
        }

        bool wrapText;

        public bool WrapText
        {
            get { return wrapText; }
            set { wrapText = value; }
        }

    }
}
