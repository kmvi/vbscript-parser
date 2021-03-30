using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class DoStatement : Statement
    {
        public DoStatement(LoopType loopType, ConditionTestType testType, Expression condition)
        {
            LoopType = loopType;
            TestType = testType;
            Condition = condition;
        }

        public LoopType LoopType { get; }
        public ConditionTestType TestType { get; }
        public Expression Condition { get; }
        public List<Statement> Body { get; } = new();
    }

    public enum ConditionTestType
    {
        None,
        PreTest,
        PostTest,
    }

    public enum LoopType
    {
        None,
        While,
        Until,
    }
}
