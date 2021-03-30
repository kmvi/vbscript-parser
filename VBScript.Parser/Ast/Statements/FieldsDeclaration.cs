using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class FieldsDeclaration : Statement
    {
        public FieldsDeclaration(FieldAccessModifier modifier)
        {
            Modifier = modifier;
        }

        public FieldAccessModifier Modifier { get; }
        public List<FieldDeclaration> Fields { get; } = new();
    }

    public enum FieldAccessModifier
    {
        Private,
        Public,
    }
}
