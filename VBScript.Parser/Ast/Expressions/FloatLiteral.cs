using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("{Value}")]
    public class FloatLiteral : LiteralExpression
    {
        public FloatLiteral(double value)
        {
            Value = value;
        }
        
        public double Value { get; }
    }
}
