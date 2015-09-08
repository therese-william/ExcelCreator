using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace ExcelCreator
{
    public class EmptyExcelException : Exception
    {
        public EmptyExcelException()
        : base() { }
    
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
