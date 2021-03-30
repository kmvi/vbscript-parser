using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class SelectStatement : Statement
    {
        public SelectStatement(Expression condition)
        {
            Condition = condition ?? throw new ArgumentNullException(nameof(condition));
        }

        public Expression Condition { get; }
        public List<CaseStatement> Cases { get; } = new();
    }
}
