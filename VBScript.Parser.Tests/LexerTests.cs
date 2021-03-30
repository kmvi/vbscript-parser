using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace VBScript.Parser.Tests
{
    public class LexerTests
    {
        [Fact]
        public void Ctor_Null_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => new Lexer(null!));
        }

        [Theory]
        [InlineData("")]
        [InlineData("        ")]
        [InlineData("    \t   \t ")]
        [InlineData(" _   \t  \r  _ \t \r\n   _ \n \t  ")]
        public void SkipWhitespaces_Whitespaces_ShouldStopAtEnd(string code)
        {
            var lexer = new Lexer(code);
            lexer.SkipWhitespaces();
            Assert.Equal(code.Length, lexer.Index);
        }

        [Theory]
        [InlineData("\n", 0)]
        [InlineData("   \n", 3)]
        [InlineData("  \t \n", 4)]
        [InlineData("  \t \r\n", 4)]
        [InlineData("  \t :", 4)]
        public void SkipWhitespaces_NewLines_ShouldStopOnNewLine(string code, int pos)
        {
            var lexer = new Lexer(code);
            lexer.SkipWhitespaces();
            Assert.Equal(pos, lexer.Index);
        }

        [Theory]
        [InlineData("")]
        [InlineData("        ")]
        [InlineData("    \t   \t ")]
        [InlineData(" _  \r   \t   \t ")]
        [InlineData(" _  \r   \t  _\n \t ")]
        [InlineData(" _  \r   \t  _ \t\r\n \t ")]
        [InlineData("_\n_\n")]
        public void NextToken_Whitespaces_ShouldReturnEofToken(string code)
        {
            var lexer = new Lexer(code);
            Assert.IsType<EofToken>(lexer.NextToken());
        }

        [Theory]
        [InlineData("\"\"", "", 0, 1)]
        [InlineData("\"test\"", "test", 0, 5)]
        [InlineData("\"Some \"\"text\"\"\"", "Some \"text\"", 0, 14)]
        public void NextToken_ValidStringLiteral_ShouldReturnStringLiteralToken(
            string code, string expected, int start, int end)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<StringLiteralToken>(lexer.NextToken());
            Assert.Equal(expected, token!.Value);
            Assert.Equal(start, token.Start);
            Assert.Equal(end, token.End);
        }

        [Theory]
        [InlineData("\"", 1, 1)]
        [InlineData("\"abc", 1, 4)]
        [InlineData("\"abc\r\"", 1, 4)]
        [InlineData("\"abc\n\"", 1, 4)]
        public void NextToken_InvalidStringLiteral_ThrowsUnterminatedStringConstant(
            string code, int line, int pos)
        {
            var lexer = new Lexer(code);
            var ex = Assert.Throws<VBSyntaxErrorException>(() => lexer.NextToken());
            Assert.Equal(VBSyntaxErrorCode.UnterminatedStringConstant, ex.Code);
            Assert.Equal(line, ex.Line);
            Assert.Equal(pos, ex.Position);
        }

        [Theory]
        [InlineData("0", 0)]
        [InlineData("1234", 1234)]
        [InlineData("53456", 53456)]
        [InlineData("2147483647", 2147483647)]
        public void NextToken_ValidDecIntLiteral_ShouldReturnIntLiteralToken(
            string code, int expected)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<DecIntegerLiteralToken>(lexer.NextToken());
            Assert.Equal(expected, token!.Value);
        }

        [Theory]
        [InlineData("2147483648", 2147483648.0)]
        [InlineData("4167923142", 4167923142.0)]
        public void NextToken_ValidLongDecIntLiteral_ShouldReturnFloatLiteralToken(
            string code, double expected)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<FloatLiteralToken>(lexer.NextToken());
            Assert.Equal(expected, token!.Value, 9);
        }

        [Theory]
        [InlineData("0eaf")]
        [InlineData("12cc34y")]
        [InlineData("53456nhn")]
        public void NextToken_InvalidDecIntLiteral_ThrowsException(
            string code)
        {
            var lexer = new Lexer(code);
            var ex = Assert.Throws<VBSyntaxErrorException>(() => lexer.NextToken());
            Assert.Equal(VBSyntaxErrorCode.ExpectedEndOfStatement, ex.Code);
            //Assert.Equal(line, ex.Line);
            //Assert.Equal(pos, ex.Position);
        }

        [Theory]
        [InlineData("&h0", 0x0)]
        [InlineData("&H0", 0x0)]
        [InlineData("&H9", 0x9)]
        [InlineData("&h100", 0x100)]
        [InlineData("&hff", 0xff)]
        [InlineData("&hfBc8E", 0xfBc8E)]
        public void NextToken_ValidHexIntLiteral_ShouldReturnIntLiteralToken(
            string code, int expected)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<HexIntegerLiteralToken>(lexer.NextToken());
            Assert.Equal(expected, token!.Value);
        }

        [Theory]
        [InlineData("&H80000000", -2147483648.0)]
        [InlineData("&hF86D75C6", -127044154.0)]
        public void NextToken_ValidLongHexIntLiteral_ShouldReturnFloatLiteralToken(
            string code, double expected)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<HexIntegerLiteralToken>(lexer.NextToken());
            Assert.Equal(expected, token!.Value, 9);
        }

        [Theory]
        [InlineData("&h0bk")]
        [InlineData("&H04r")]
        [InlineData("&H9fn56")]
        public void NextToken_InvalidHexIntLiteral_ThrowsException(
            string code)
        {
            var lexer = new Lexer(code);
            var ex = Assert.Throws<VBSyntaxErrorException>(() => lexer.NextToken());
            Assert.Equal(VBSyntaxErrorCode.ExpectedEndOfStatement, ex.Code);
        }

        [Theory]
        [InlineData("&0", 00)]
        [InlineData("&7", 07)]
        [InlineData("&100", 64)]
        [InlineData("&7654321", 2054353)]
        [InlineData("&01234567", 342391)]
        public void NextToken_ValidOctIntLiteral_ShouldReturnIntLiteralToken(
            string code, int expected)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<OctIntegerLiteralToken>(lexer.NextToken());
            Assert.Equal(expected, token!.Value);
        }

        [Theory]
        [InlineData("&20000000000", -2147483648.0)]
        [InlineData("&37033272706", -127044154.0)]
        public void NextToken_ValidLongOctIntLiteral_ShouldReturnFloatLiteralToken(
            string code, double expected)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<OctIntegerLiteralToken>(lexer.NextToken());
            Assert.Equal(expected, token.Value, 9);
        }

        [Theory]
        [InlineData("&9")]
        [InlineData("&7842943")]
        [InlineData("&10a0")]
        [InlineData("&76k54321")]
        [InlineData("&01234567y")]
        public void NextToken_InvalidOctIntLiteral_ThrowsException(string code)
        {
            var lexer = new Lexer(code);
            var ex = Assert.Throws<VBSyntaxErrorException>(() => lexer.NextToken());
            Assert.True(VBSyntaxErrorCode.ExpectedEndOfStatement == ex.Code ||
                VBSyntaxErrorCode.SyntaxError == ex.Code);
        }

        [Theory]
        [InlineData("  \t  ' test  ", " test  ")]
        [InlineData("  \t  rem test  ", " test  ")]
        [InlineData("  \t  REM test  ", " test  ")]
        [InlineData(" \t \t  REM test  ", " test  ")]
        [InlineData("REM test 5:43 \r\n", " test 5:43 ")]
        [InlineData(" ' test 5:43 \r\n", " test 5:43 ")]
        [InlineData(" _ \n_\n_\r_\r\n   \t ' test 5:43 \r\n", " test 5:43 ")]
        public void NextToken_ValidComment_ShouldReturnCommentToken(string code, string expected)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<CommentToken>(lexer.NextToken());
            Assert.Equal(expected, token!.Comment);
        }

        [Theory]
        [InlineData("\r")]
        [InlineData("\n")]
        [InlineData(":")]
        [InlineData("    \t \t \r\r\r  \n   \t")]
        [InlineData("  :  \t \t \r:\r\r  \n   \t")]
        [InlineData("    \t \t \r\r\r::\n:   :  \n   \t")]
        [InlineData("  _ \r  \t \t \r\r\r::\n:   :  \n   \t")]
        public void NextToken_NewLines_ShouldReturnLineTerminationToken(string code)
        {
            var lexer = new Lexer(code);
            Assert.IsAssignableFrom<LineTerminationToken>(lexer.NextToken());
            Assert.IsType<EofToken>(lexer.NextToken());
        }

        [Theory]
        [InlineData("0.", .0d)]
        [InlineData("000.", .0d)]
        [InlineData(".0", .0d)]
        [InlineData(".000", .0d)]
        [InlineData(".0e0", .0d)]
        [InlineData(".000e000", .0d)]
        [InlineData("00.000e0000", .0d)]
        [InlineData(".0e-0", .0d)]
        [InlineData(".0e+0", .0d)]
        [InlineData("00.000E-0000", .0d)]
        [InlineData(".000e+000", .0d)]
        [InlineData("00.000e+0000", .0d)]
        [InlineData("245.", 245.0d)]
        [InlineData("0345.", 345.0d)]
        [InlineData(".04023", .04023d)]
        [InlineData(".2310e34", .2310e34d)]
        [InlineData("66.235E13", 66.235E13d)]
        [InlineData("54.0456e-4", 54.0456e-4d)]
        [InlineData("02.123E+2", 02.123E+2d)]
        public void NextToken_ValidFloatLiteral_ShouldReturnFloatLiteralToken(
            string code, double expected)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<FloatLiteralToken>(lexer.NextToken());
            Assert.Equal(expected, token!.Value, 9);
        }

        [Theory]
        [InlineData("<>", Punctuation.NotEqual)]
        [InlineData("><", Punctuation.NotEqual)]
        [InlineData("<=", Punctuation.LessOrEqual)]
        [InlineData(">=", Punctuation.GreaterOrEqual)]
        [InlineData("=<", Punctuation.LessOrEqual)]
        [InlineData("=>", Punctuation.GreaterOrEqual)]
        public void NextToken_ValidTwoCharPunctuation_ShourdReturnValidPunctuationToken(
            string code, Punctuation type)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<PunctuationToken>(lexer.NextToken());
            Assert.Equal(type, token.Type);
        }

        [Theory]
        [InlineData("    abc123_4ABC")]
        [InlineData("abc123_4ABC")]
        [InlineData("Nothing_notLiteral")]
        [InlineData("TrueNotLiteral")]
        [InlineData("Falsee")]
        [InlineData("Empty1")]
        [InlineData("Null_Null")]
        public void NextToken_ValidIdentifier_ReturnsIdentifier(string code)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<IdentifierToken>(lexer.NextToken());
            Assert.Equal(code.Trim(), token.Name);
        }

        [Theory]
        [InlineData("\t  [abc1234ABC]")]
        [InlineData("[#^&%^&()@[]")]
        public void NextToken_ValidExtendedIdentifier_ReturnsIdentifier(string code)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<ExtendedIdentifierToken>(lexer.NextToken());
            Assert.Equal(code.Trim(), token.Name);
        }

        [Theory]
        [InlineData("#2019-01-02#", "2019-01-02T00:00:00")]
        [InlineData("#03/08/2018#", "2018-03-08T00:00:00")]
        [InlineData("# \t 2019-03-12 #", "2019-03-12T00:00:00")]
        [InlineData("# \t Jul 7, 1992 #", "1992-07-07T00:00:00")]
        public void NextToken_ValidDateLiteral_ReturnsDate(string code, string dt)
        {
            var lexer = new Lexer(code);
            var token = Assert.IsType<DateLiteralToken>(lexer.NextToken());
            Assert.Equal(DateTime.Parse(dt), token.Value);
        }

        [Theory]
        [InlineData("true")]
        [InlineData("\t  \t   True")]
        [InlineData("TrUe")]
        public void NextToken_TrueLiteral(string code)
        {
            var lexer = new Lexer(code);
            Assert.IsType<TrueLiteralToken>(lexer.NextToken());
        }

        [Theory]
        [InlineData("false")]
        [InlineData("   \t  False")]
        [InlineData("fAlsE")]
        public void NextToken_FalseLiteral(string code)
        {
            var lexer = new Lexer(code);
            Assert.IsType<FalseLiteralToken>(lexer.NextToken());
        }

        [Theory]
        [InlineData("null")]
        [InlineData("\t  \t   Null")]
        [InlineData("NuLL")]
        public void NextToken_NullLiteral(string code)
        {
            var lexer = new Lexer(code);
            Assert.IsType<NullLiteralToken>(lexer.NextToken());
        }

        [Theory]
        [InlineData("empty")]
        [InlineData(" \t  \t   Empty")]
        [InlineData("EmPtY")]
        public void NextToken_EmptyLiteral(string code)
        {
            var lexer = new Lexer(code);
            Assert.IsType<EmptyLiteralToken>(lexer.NextToken());
        }

        [Theory]
        [InlineData("nothing")]
        [InlineData(" \t  \t   Nothing")]
        [InlineData("NoThIng")]
        public void NextToken_NothingLiteral(string code)
        {
            var lexer = new Lexer(code);
            Assert.IsType<NothingLiteralToken>(lexer.NextToken());
        }
    }
}
