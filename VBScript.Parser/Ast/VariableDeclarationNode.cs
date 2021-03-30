using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public abstract class VariableDeclarationNode : Node
    {
        protected VariableDeclarationNode(Identifier id, bool isDynamicArray)
        {
            Identifier = id ?? throw new ArgumentNullException(nameof(id));
            IsDynamicArray = isDynamicArray;
        }

        public Identifier Identifier { get; }
        public bool IsDynamicArray { get; }
        public List<int> ArrayDims { get; } = new();
    }
}
