using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class ClassDeclaration : Statement
    {
        public ClassDeclaration(Identifier id)
        {
            Identifier = id ?? throw new ArgumentNullException(nameof(id));
        }

        public Identifier Identifier { get; }
        public List<Statement> Members { get; } = new();
    }
}
