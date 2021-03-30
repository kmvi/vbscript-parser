using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("Function {Identifier.Name}")]
    public class FunctionDeclaration : ProcedureDeclaration
    {
        public FunctionDeclaration(MethodAccessModifier modifier, Identifier id, Statement body)
            : base(modifier, id, body)
        {
        }
    }
}
