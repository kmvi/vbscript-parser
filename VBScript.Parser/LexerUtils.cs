using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace VBScript.Parser
{
    static class LexerUtils
    {
        private static readonly Dictionary<string, Keyword> _keywords = new(StringComparer.OrdinalIgnoreCase)
        {
            ["and"] = Keyword.And,
            ["byref"] = Keyword.ByRef,
            ["byval"] = Keyword.ByVal,
            ["call"] = Keyword.Call,
            ["case"] = Keyword.Case,
            ["class"] = Keyword.Class,
            ["const"] = Keyword.Const,
            ["dim"] = Keyword.Dim,
            ["do"] = Keyword.Do,
            ["each"] = Keyword.Each,
            ["else"] = Keyword.Else,
            ["elseif"] = Keyword.ElseIf,
            ["end"] = Keyword.End,
            ["eqv"] = Keyword.Eqv,
            ["exit"] = Keyword.Exit,
            ["for"] = Keyword.For,
            ["function"] = Keyword.Function,
            ["get"] = Keyword.Get,
            ["goto"] = Keyword.Goto,
            ["if"] = Keyword.If,
            ["imp"] = Keyword.Imp,
            ["in"] = Keyword.In,
            ["is"] = Keyword.Is,
            ["let"] = Keyword.Let,
            ["loop"] = Keyword.Loop,
            ["mod"] = Keyword.Mod,
            ["new"] = Keyword.New,
            ["next"] = Keyword.Next,
            ["not"] = Keyword.Not,
            ["on"] = Keyword.On,
            ["option"] = Keyword.Option,
            ["or"] = Keyword.Or,
            ["preserve"] = Keyword.Preserve,
            ["private"] = Keyword.Private,
            ["public"] = Keyword.Public,
            ["redim"] = Keyword.ReDim,
            ["resume"] = Keyword.Resume,
            ["select"] = Keyword.Select,
            ["set"] = Keyword.Set,
            ["sub"] = Keyword.Sub,
            ["then"] = Keyword.Then,
            ["to"] = Keyword.To,
            ["until"] = Keyword.Until,
            ["wend"] = Keyword.WEnd,
            ["while"] = Keyword.While,
            ["with"] = Keyword.With,
            ["xor"] = Keyword.Xor,
        };

        private static readonly Dictionary<string, Keyword> _keywordAsIdentifiers = new(StringComparer.OrdinalIgnoreCase)
        {
            ["default"] = Keyword.Default,
            ["erase"] = Keyword.Erase,
            ["error"] = Keyword.Error,
            ["explicit"] = Keyword.Explicit,
            ["property"] = Keyword.Property,
            ["step"] = Keyword.Step,
        };

        public static bool IsKeyword(string s)
            => _keywords.ContainsKey(s);

        public static Keyword GetKeyword(string s)
            => _keywords[s];

        public static bool IsKeywordAsIdentifier(string s)
            => _keywordAsIdentifiers.ContainsKey(s);

        public static Keyword GetKeywordAsIdentifier(string s)
            => _keywordAsIdentifiers[s];

        public static DateTime GetDate(string s)
        {
            // TODO:
            return DateTime.Parse(s);
        }
    }
}
