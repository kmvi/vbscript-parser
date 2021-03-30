using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class ForEachStatement : Statement
    {
        public ForEachStatement(Identifier id, Expression @in)
        {
            Identifier = id ?? throw new ArgumentNullException(nameof(id));
            In = @in ?? throw new ArgumentNullException(nameof(@in));
        }

        public Identifier Identifier { get; }
        public Expression In { get; }
        public List<Statement> Body { get; } = new();
    }
}
