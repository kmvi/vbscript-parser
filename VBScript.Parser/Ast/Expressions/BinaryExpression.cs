using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class BinaryExpression : Expression
    {
        public BinaryExpression(BinaryOperation op, Expression left, Expression right)
        {
            Operation = op;
            Left = left;
            Right = right;
        }

        public BinaryOperation Operation { get; }
        public Expression Left { get; }
        public Expression Right { get; }
    }

    public enum BinaryOperation
    {
        Exponentiation,
        Multiplication,
        Division,
        IntDivision,
        Addition,
        Subtraction,
        Concatenation,
        Mod,
        Is,
        And,
        Or,
        Xor,
        Eqv,
        Imp,
        Equal,
        NotEqual,
        Less,
        Greater,
        LessOrEqual,
        GreaterOrEqual,
    }
}
