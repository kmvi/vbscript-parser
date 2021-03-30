using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class PropertyLetDeclaration : PropertyDeclaration
    {
        public PropertyLetDeclaration(MethodAccessModifier modifier, Identifier id)
            : base(modifier, id)
        {
        }
    }
}
