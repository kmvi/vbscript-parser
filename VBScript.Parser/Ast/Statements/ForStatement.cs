using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class ForStatement : Statement
    {
        public ForStatement(Identifier id, Expression from, Expression to, Expression? step)
        {
            Identifier = id ?? throw new ArgumentNullException(nameof(id));
            From = from;
            To = to;
            Step = step;
        }

        public Identifier Identifier { get; }
        public Expression From { get; }
        public Expression To { get; }
        public Expression? Step { get; }
        public List<Statement> Body { get; } = new();
    }
}
