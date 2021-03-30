using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace VBScript.Parser.Ast
{
    [DebuggerDisplay("{Text}")]
    public class Comment
    {
        public Comment(CommentType type, string text)
        {
            Type = type;
            Text = text;
        }

        public CommentType Type { get; }
        public string Text { get; }
        public Range Range { get; set; }
        public Location Location { get; set; }
    }

    public enum CommentType
    {
        Rem,
        SingleQuote
    }
}
