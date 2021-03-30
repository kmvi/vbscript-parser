using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class Program : Node
    {
        public Program(bool optionExplicit)
        {
            OptionExplicit = optionExplicit;
        }

        public bool OptionExplicit { get; }
        public List<Statement> Body { get; } = new();
        public List<Comment> Comments { get; } = new();
    }
}
