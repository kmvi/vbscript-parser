using System;
using System.Collections.Generic;
using System.Resources;
using System.Text;

namespace VBScript.Parser
{
    static class Extensions
    {
        public static char GetChar(this string str, int pos)
            => pos < 0 || pos >= str.Length ? Char.MinValue : str[pos];

        public static int GetCharCode(this string str, int pos)
            => pos < 0 || pos >= str.Length ? Char.MinValue : str[pos];

        public static string Slice(this string str, int start, int end)
        {
            var len = str.Length;
            var from = start < 0 ? Math.Max(len + start, 0) : Math.Min(start, len);
            var to = end < 0 ? Math.Max(len + end, 0) : Math.Min(end, len);
            var span = Math.Max(to - from, 0);
            var substring = str.Substring(from, span);
            return substring;
        }

        public static bool CIEquals(this string s1, string s2)
            => String.Equals(s1, s2, StringComparison.OrdinalIgnoreCase);

        public static string GetString(this ResourceManager manager, VBSyntaxErrorCode code)
            => manager.GetString(((int)code).ToString());
    }
}
