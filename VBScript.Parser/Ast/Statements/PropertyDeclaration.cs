using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public abstract class PropertyDeclaration : Statement
    {
        protected PropertyDeclaration(MethodAccessModifier modifier, Identifier id)
        {
            AccessModifier = modifier;
            Identifier = id ?? throw new ArgumentNullException(nameof(id));
        }

        public MethodAccessModifier AccessModifier { get; }
        public Identifier Identifier { get; }
        public List<Parameter> Parameters { get; } = new();
        public List<Statement> Body { get; } = new();
    }
}
