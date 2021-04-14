using System;
using System.IO;
using System.Text;

namespace VBScript.Parser.Example
{
    class Program
    {
        static void Main(string[] args)
        {
            string code = @"
' comment 1
rem comment 2

option explicit

sub s(k)
    with wscript ' comment 3
        .echo k
    end with
end sub

dim a, b(3), i

a = 234 & ""test""
s a

for i = lbound(b) to ubound(b)
    b(i) = i * 2
    call s(b(i))
next
";

            var parser = new VBScriptParser(code, new ParsingOptions { SaveComments = true });
            var program = parser.Parse();
        }
    }
}
