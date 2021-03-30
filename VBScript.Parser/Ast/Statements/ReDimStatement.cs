using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class ReDimStatement : Statement
    {
        public ReDimStatement(bool preserve)
        {
            Preserve = preserve;
        }

        public bool Preserve { get; }
        public List<ReDimDeclaration> ReDims { get; } = new();
    }
}
