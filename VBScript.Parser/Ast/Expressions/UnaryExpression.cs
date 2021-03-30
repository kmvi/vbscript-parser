using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class UnaryExpression : Expression
    {
        public UnaryExpression(UnaryOperation operation, Expression arg)
        {
            Operation = operation;
            Argument = arg ?? throw new ArgumentNullException(nameof(arg));
        }

        public UnaryOperation Operation { get; }
        public Expression Argument { get; }
    }

    public enum UnaryOperation
    {
        Plus,
        Minus,
        Not,
    }
}
