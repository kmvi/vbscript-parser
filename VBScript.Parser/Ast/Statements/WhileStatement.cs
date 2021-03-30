using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class WhileStatement : Statement
    {
        public WhileStatement(Expression condition)
        {
            Condition = condition ?? throw new ArgumentNullException(nameof(condition));
        }

        public Expression Condition { get; }
        public List<Statement> Body { get; } = new();
    }
}
