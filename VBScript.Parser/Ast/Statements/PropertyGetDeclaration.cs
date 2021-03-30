using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class PropertyGetDeclaration : PropertyDeclaration
    {
        public PropertyGetDeclaration(MethodAccessModifier modifier, Identifier id)
            : base(modifier, id)
        {
        }
    }
}
