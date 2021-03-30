using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class ElseIfStatement : Statement
    {
        public ElseIfStatement(Expression test, Statement consequent, Statement? alternate)
        {
            Test = test ?? throw new ArgumentNullException(nameof(test));
            Consequent = consequent ?? throw new ArgumentNullException(nameof(consequent));
            Alternate = alternate;
        }

        public Expression Test { get; }
        public Statement Consequent { get; }
        public Statement? Alternate { get; }
    }
}
