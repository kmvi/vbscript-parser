using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("{Identifier.Name}")]
    public class FieldDeclaration : VariableDeclarationNode
    {
        public FieldDeclaration(Identifier id, bool isDynamicArray)
            : base(id, isDynamicArray)
        {
        }
    }
}
