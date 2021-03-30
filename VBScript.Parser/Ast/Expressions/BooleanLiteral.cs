using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("{Value}")]
    public class BooleanLiteral : LiteralExpression
    {
        public BooleanLiteral(bool value)
        {
            Value = value;
        }
        
        public bool Value { get; }
    }
}
