using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class ReDimDeclaration : Node
    {
        public ReDimDeclaration(Identifier id)
        {
            Identifier = id ?? throw new ArgumentNullException(nameof(id));
        }

        public Identifier Identifier { get; }
        public List<Expression> ArrayDims { get; } = new();
    }
}
