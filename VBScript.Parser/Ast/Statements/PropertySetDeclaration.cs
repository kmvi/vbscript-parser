using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class PropertySetDeclaration : PropertyDeclaration
    {
        public PropertySetDeclaration(MethodAccessModifier modifier, Identifier id)
            : base(modifier, id)
        {
        }
    }
}
