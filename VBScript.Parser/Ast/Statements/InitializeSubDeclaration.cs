using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("Sub {Identifier.Name}")]
    public class InitializeSubDeclaration : SubDeclaration
    {
        public static readonly string Name = "Class_Initialize";

        public InitializeSubDeclaration(MethodAccessModifier modifier, Statement body)
            : base(modifier, new Identifier(Name), body)
        {
        }
    }
}
