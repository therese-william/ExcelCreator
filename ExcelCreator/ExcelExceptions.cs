using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace ExcelCreator
{
    /// <summary>
    /// EmptyExcelException is raised when excel sheet have no sheets
    /// </summary>
    public class EmptyExcelException : Exception
    {
        /// <summary>
        /// EmptyExcelException fired when excel sheet have no sheets
        /// </summary>
        public EmptyExcelException()
        : base() { }

        /// <summary>
        /// EmptyExcelException fired when excel sheet have no sheets
        /// </summary>
        /// <param name="message">message appear to user</param>
        public EmptyExcelException(string message)
            : base(message) { }
    
        public EmptyExcelException(string format, params object[] args)
            : base(string.Format(format, args)) { }
    
        public EmptyExcelException(string message, Exception innerException)
            : base(message, innerException) { }
    
        public EmptyExcelException(string format, Exception innerException, params object[] args)
            : base(string.Format(format, args), innerException) { }

        protected EmptyExcelException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }
    /// <summary>
    /// InvalidSheetException is raised when there is something invalid in the sheet. i.e. number of values in supplied rows don't match number of columns
    /// </summary>
    public class InvalidSheetException : Exception
    {
        public InvalidSheetException()
            : base() { }

        public InvalidSheetException(string message)
            : base(message) { }

        public InvalidSheetException(string format, params object[] args)
            : base(string.Format(format, args)) { }

        public InvalidSheetException(string message, Exception innerException)
            : base(message, innerException) { }

        public InvalidSheetException(string format, Exception innerException, params object[] args)
            : base(string.Format(format, args), innerException) { }

        protected InvalidSheetException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }
}
