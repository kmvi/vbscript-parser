using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class CallSubStatement : Statement
    {
        public CallSubStatement(Expression callee)
        {
            Callee = callee ?? throw new ArgumentNullException(nameof(callee));
        }

        public Expression Callee { get; }
        public List<Expression> Arguments { get; } = new();
    }
}
