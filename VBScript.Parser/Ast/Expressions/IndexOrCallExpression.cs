using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class IndexOrCallExpression : Expression
    {
        public IndexOrCallExpression(Expression obj)
        {
            Object = obj ?? throw new ArgumentNullException(nameof(obj));
        }

        public List<Expression> Indexes { get; } = new();
        public Expression Object { get; }
    }
}
