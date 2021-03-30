using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class NewExpression : Expression
    {
        public NewExpression(Expression arg)
        {
            Argument = arg ?? throw new ArgumentNullException();
        }

        public Expression Argument { get; }
    }
}
