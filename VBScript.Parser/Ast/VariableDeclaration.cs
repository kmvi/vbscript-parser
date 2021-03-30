using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("{Identifier.Name}")]
    public class VariableDeclaration : VariableDeclarationNode
    {
        public VariableDeclaration(Identifier id, bool isDynamicArray)
            : base(id, isDynamicArray)
        {
        }
    }
}
