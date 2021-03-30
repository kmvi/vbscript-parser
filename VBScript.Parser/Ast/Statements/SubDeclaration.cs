using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("Sub {Identifier.Name}")]
    public class SubDeclaration : ProcedureDeclaration
    {
        public SubDeclaration(MethodAccessModifier modifier, Identifier id, Statement body)
            : base(modifier, id, body)
        {
        }
    }
}
