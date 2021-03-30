using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class WithStatement : Statement
    {
        public WithStatement(Expression expr)
        {
            Expression = expr ?? throw new ArgumentNullException(nameof(expr));
        }

        public Expression Expression { get; }
        public List<Statement> Body { get; } = new();
    }
}
