using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class EraseStatement : Statement
    {
        public EraseStatement(Identifier id)
        {
            Identifier = id ?? throw new ArgumentNullException(nameof(id));
        }

        public Identifier Identifier { get; }
    }
}
