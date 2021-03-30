using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace VBScript.Parser.Tests
{
    public class ParserTests
    {
        [Fact]
        public void Ctor_NullArguments_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => new Parser(null!));
            Assert.Throws<ArgumentNullException>(() => new Parser(null!, new()));
            Assert.Throws<ArgumentNullException>(() => new Parser("", null!));
            Assert.Throws<ArgumentNullException>(() => new Parser(null!, null!));
        }

        [Theory]
        [InlineData("")]
        [InlineData("  :  \t   \t \n \r \r\n \t ")]
        [InlineData("  :  \t  ' test comment \t \n \r \r\n \t ")]
        public void Parse_EmptyString_ReturnsEmptyAst(string code)
        {
            var prg = new Parser(code).Parse();
            
            Assert.NotNull(prg);
            Assert.False(prg.OptionExplicit);
            Assert.Empty(prg.Comments);
            Assert.Empty(prg.Body);
            Assert.Equal(0, prg.Range.Start);
        }

        [Theory]
        [InlineData("rem some text")]
        [InlineData("' some text")]
        [InlineData("  \t:::\n  \r\n  \t :' some text   \t \r\n:")]
        public void Parse_EmptyStringWithComments_ReturnsEmptyAstWithComments(string code)
        {
            var prg = new Parser(code, new ParsingOptions { SaveComments = true }).Parse();

            Assert.NotNull(prg);
            Assert.False(prg.OptionExplicit);
            Assert.Single(prg.Comments);
            Assert.Empty(prg.Body);
            Assert.Equal(0, prg.Range.Start);
        }

        [Theory]
        [InlineData("option explicit")]
        [InlineData("option _\nexplicit")]
        [InlineData("'test\noption explicit  'test ")]
        public void Parse_OptionExplicitString_ReturnsAstWithEmptyBody(string code)
        {
            var prg = new Parser(code).Parse();

            Assert.NotNull(prg);
            Assert.True(prg.OptionExplicit);
            Assert.Empty(prg.Comments);
            Assert.Empty(prg.Body);
            Assert.Equal(0, prg.Range.Start);
        }
    }
}
