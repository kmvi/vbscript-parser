using System;

namespace VBScript.Parser
{
    public abstract class Token
    {
        public int Start { get; set; }
        public int End { get; set; }
        public int LineNumber { get; set; }
        public int LineStart { get; set; }
    }

    public class EofToken : Token
    {
    }

    public class LineTerminationToken : Token
    {
    }

    public class ColonLineTerminationToken : LineTerminationToken
    {
    }

    public class CommentToken : Token
    {
        public string Comment { get; set; } = "";
        public bool IsRem { get; set; }
    }

    public abstract class LiteralToken : Token
    {

    }

    public class StringLiteralToken : LiteralToken
    {
        public string Value { get; set; } = "";
    }

    public class DecIntegerLiteralToken : LiteralToken
    {
        public int Value { get; set; }
    }

    public class HexIntegerLiteralToken : DecIntegerLiteralToken
    {
    }

    public class OctIntegerLiteralToken : DecIntegerLiteralToken
    {
    }

    public class DateLiteralToken : LiteralToken
    {
        public DateTime Value { get; set; }
    }

    public class FloatLiteralToken : LiteralToken
    {
        public double Value { get; set; }
    }

    public class TrueLiteralToken : LiteralToken
    {
    }

    public class FalseLiteralToken : LiteralToken
    {
    }

    public class NullLiteralToken : LiteralToken
    {
    }

    public class NothingLiteralToken : LiteralToken
    {
    }

    public class EmptyLiteralToken : LiteralToken
    {
    }

    public class IdentifierToken : Token
    {
        public string Name { get; set; } = "";
        public override string ToString() => Name;
    }

    public class KeywordToken : Token
    {
        public Keyword Keyword { get; set; }
        public string Name { get; set; } = "";
        public override string ToString() => Name;
    }

    public class KeywordOrIdentifierToken : Token
    {
        public string Name { get; set; } = "";
        public override string ToString() => Name;
        public Keyword Keyword { get; set; }
    }

    public class ExtendedIdentifierToken : IdentifierToken
    {
        public override string ToString() => "[" + Name + "]";
    }

    public class PunctuationToken : Token
    {
        public Punctuation Type { get; set; }
    }

    internal class InvalidToken : Token
    {
    }
}
