using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class AssignmentStatement : Statement
    {
        public AssignmentStatement(Expression left, Expression right, bool set)
        {
            Left = left ?? throw new ArgumentNullException(nameof(left));
            Right = right ?? throw new ArgumentNullException(nameof(right));
            Set = set;
        }

        public bool Set { get; }
        public Expression Left { get; }
        public Expression Right { get; }
    }
}
