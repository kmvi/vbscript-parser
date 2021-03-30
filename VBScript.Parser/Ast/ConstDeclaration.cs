using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class ConstDeclaration : Node
    {
        public ConstDeclaration(Identifier id, Expression init)
        {
            Identifier = id ?? throw new ArgumentNullException(nameof(id));
            Init = init ?? throw new ArgumentNullException(nameof(init));
        }

        public Identifier Identifier { get; }
        public Expression Init { get; }
    }
}
