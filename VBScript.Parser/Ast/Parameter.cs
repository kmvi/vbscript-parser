using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class Parameter : Node
    {
        public Parameter(Identifier id, ParameterModifier modifier, bool parentheses)
        {
            Modifier = modifier;
            Parentheses = parentheses;
            Identifier = id ?? throw new ArgumentNullException(nameof(id));
        }

        public ParameterModifier Modifier { get; }
        public bool Parentheses { get; }
        public Identifier Identifier { get; }
    }

    public enum ParameterModifier
    {
        None, ByRef, ByVal
    }
}
