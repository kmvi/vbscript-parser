using System;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser
{
    static class CharUtils
    {
        public static bool IsLineTerminator(char c)
            => c == '\n' ||
               c == '\r' ||
               c == ':';

        public static bool IsNewLine(char c)
            => c == '\n' ||
               c == '\r';

        public static bool IsDecDigit(char c)
            => c >= '0' && c <= '9';

        public static bool IsHexDigit(char c)
            => IsDecDigit(c) ||
               (c >= 'a' && c <= 'f') ||
               (c >= 'A' && c <= 'F');

        public static bool IsOctDigit(char c)
            => c >= '0' && c <= '7';

        public static bool IsIdentifierStart(char c)
            => (c >= 'A' && c <= 'Z') ||
               (c >= 'a' && c <= 'z');

        public static bool IsIdentifier(char c)
            => IsIdentifierStart(c) ||
               IsDecDigit(c) ||
               c == '_';

        public static bool IsWhiteSpace(char c)
            => c == 0x20 || c == 0x09 ||
               c == 0x0B || c == 0x0C;

        public static bool IsExtendedIdentifier(char c)
            => !IsNewLine(c) && c != ']' && c >= 0 && c <= 0xff;

        public static bool Equals(char a, char b)
            => Char.ToUpperInvariant(a) == Char.ToUpperInvariant(b);
    }
}
