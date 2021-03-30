using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    public abstract class ProcedureDeclaration : Statement
    {
        protected ProcedureDeclaration(MethodAccessModifier modifier, Identifier id, Statement body)
        {
            AccessModifier = modifier;
            Identifier = id ?? throw new ArgumentNullException(nameof(id));
            Body = body ?? throw new ArgumentNullException(nameof(body));
        }

        public Identifier Identifier { get; }
        public MethodAccessModifier AccessModifier { get; }
        public List<Parameter> Parameters { get; } = new();
        public Statement Body { get; }
    }

    public enum MethodAccessModifier
    {
        None,
        Public,
        Private,
        PublicDefault,
    }
}
