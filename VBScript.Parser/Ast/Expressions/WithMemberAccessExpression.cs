using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay(".{Property.Name}")]
    public class WithMemberAccessExpression : Expression
    {
        public WithMemberAccessExpression(Identifier property)
        {
            Property = property ?? throw new ArgumentNullException(nameof(property));
        }

        public Identifier Property { get; }
    }
}
