using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("{Value}")]
    public class DateLiteral : LiteralExpression
    {
        public DateLiteral(DateTime value)
        {
            Value = value;
        }
        
        public DateTime Value { get; }
    }
}
