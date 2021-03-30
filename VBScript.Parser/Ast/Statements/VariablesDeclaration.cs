using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class VariablesDeclaration : Statement
    {
        public List<VariableDeclaration> Variables { get; } = new();
    }
}
