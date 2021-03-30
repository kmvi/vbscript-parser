using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class CaseStatement : Statement
    {
        public CaseStatement()
        {

        }

        public List<Expression> Values { get; } = new();
        public List<Statement> Body { get; } = new();
    }
}
