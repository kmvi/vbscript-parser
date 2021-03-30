using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("Sub {Identifier.Name}")]
    public class TerminateSubDeclaration : SubDeclaration
    {
        public static readonly string Name = "Class_Terminate";

        public TerminateSubDeclaration(MethodAccessModifier modifier, Statement body)
            : base(modifier, new Identifier("Class_Terminate"), body)
        {
        }
    }
}
