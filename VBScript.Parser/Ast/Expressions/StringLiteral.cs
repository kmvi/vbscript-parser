using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("{Value}")]
    public class StringLiteral : LiteralExpression
    {
        public StringLiteral(string value)
        {
            Value = value;
        }
        
        public string Value { get; }
    }
}
