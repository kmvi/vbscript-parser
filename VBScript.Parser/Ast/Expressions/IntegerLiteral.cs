using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("{Value}")]
    public class IntegerLiteral : LiteralExpression
    {
        public IntegerLiteral(int value)
        {
            Value = value;
        }
        
        public int Value { get; }
    }
}
