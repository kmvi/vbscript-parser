using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class CallStatement : Statement
    {
        public CallStatement(Expression callee)
        {
            Callee = callee ?? throw new ArgumentNullException(nameof(callee));
        }

        public Expression Callee { get; }
    }
}
