using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class ConstsDeclaration : Statement
    {
        public ConstsDeclaration(MemberAccessModifier modifier)
        {
            Modifier = modifier;
        }

        public MemberAccessModifier Modifier { get; }
        public List<ConstDeclaration> Declarations { get; } = new();
    }

    public enum MemberAccessModifier
    {
        None,
        Public,
        Private,
    }
}
