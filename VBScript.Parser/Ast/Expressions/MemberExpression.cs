using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class MemberExpression : Expression
    {
        public MemberExpression(Expression obj, Identifier property)
        {
            Object = obj ?? throw new ArgumentNullException(nameof(obj));
            Property = property ?? throw new ArgumentNullException(nameof(property));
        }

        public Expression Object { get; }
        public Identifier Property { get; }
    }
}
