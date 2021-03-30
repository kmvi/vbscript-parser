using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace VBScript.Parser
{
    public class Lexer
    {
        private readonly int _length;
        private readonly StringBuilder _sb = new StringBuilder();

        public Lexer(string code)
        {
            Code = code ?? throw new ArgumentNullException(nameof(code));
            _length = Code.Length;
            CurrentLine = Code.Length == 0 ? 0 : 1;
        }

        public string Code { get; }
        public int Index { get; private set; }
        public int CurrentLine { get; private set; }
        public int CurrentLineStart { get; private set; }
        public int LineIndex => Index - CurrentLineStart;

        public Token NextToken()
        {
            SkipWhitespaces();
            
            if (IsEof())
            {
                return new EofToken
                {
                    Start = Index,
                    End = Index,
                    LineNumber = CurrentLine,
                    LineStart = CurrentLineStart,
                };
            }

            char c = Code.GetChar(Index);
            char next = Code.GetChar(Index + 1);

            if (CharUtils.IsLineTerminator(c))
                return NextLineTermination();

            var comment = NextComment();
            if (comment != null)
                return comment;

            if (CharUtils.IsIdentifierStart(c))
                return NextIdentifier();

            if (c == '"')
                return NextStringLiteral();

            if (c == '.')
            {
                if (CharUtils.IsDecDigit(next))
                {
                    return NextNumericLiteral();
                }
                return NextPunctuation();
            }

            if (CharUtils.IsDecDigit(c))
                return NextNumericLiteral();

            if (c == '&')
            {
                if (CharUtils.Equals(next, 'h') || CharUtils.IsDecDigit(next))
                {
                    return NextNumericLiteral();
                }
                return NextPunctuation();
            }

            if (c == '#')
            {
                return NextDateLiteral();
            }

            if (c == '[')
            {
                return NextExtendedIdentifier();
            }

            return NextPunctuation();
        }

        public void SkipWhitespaces()
        {
            void SkipWSOnly()
            {
                char c = Code.GetChar(Index);
                while (CharUtils.IsWhiteSpace(c))
                    c = Code.GetChar(++Index);
            }

            while (!IsEof())
            {
                SkipWSOnly();
                char c = Code.GetChar(Index);
                if (c == '_')
                {
                    ++Index;
                    SkipWSOnly();
                    c = Code.GetChar(Index);
                    if (CharUtils.IsNewLine(c))
                        SkipNewline();
                    else
                        throw VBSyntaxError(VBSyntaxErrorCode.InvalidCharacter);
                }
                else
                {
                    break;
                }
            }
        }

        public IEnumerable<Token> AsSequence()
        {
            while (!IsEof())
            {
                yield return NextToken();
            }
        }

        public void Reset()
        {
            Index = 0;
            CurrentLine = Code.Length == 0 ? 0 : 1;
            CurrentLineStart = 0;
            _sb.Clear();
        }

        private Token NextExtendedIdentifier()
        {
            int start = Index++;
            while (!IsEof())
            {
                char c = Code.GetChar(Index);
                if (CharUtils.IsExtendedIdentifier(c))
                {
                    ++Index;
                }
                else
                {
                    break;
                }
            }

            if (Code.GetChar(Index) != ']')
            {
                throw VBSyntaxError(VBSyntaxErrorCode.ExpectedRBracket);
            }

            return new ExtendedIdentifierToken
            {
                Start = start,
                End = ++Index,
                LineNumber = CurrentLine,
                LineStart = CurrentLineStart,
                Name = Code.Slice(start, Index),
            };
        }

        private Token NextDateLiteral()
        {
            int start = Index++;
            var str = GetStringBuilder();

            while (!IsEof())
            {
                char c = Code.GetChar(Index);
                if (c == '#' || CharUtils.IsNewLine(c))
                {
                    break;
                }
                else
                {
                    str.Append(c);
                    ++Index;
                }
            }

            if (Code.GetChar(Index) != '#' || str.Length == 0)
            {
                throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError);
            }

            var date = LexerUtils.GetDate(str.ToString());

            return new DateLiteralToken
            {
                Start = start,
                End = ++Index,
                LineNumber = CurrentLine,
                LineStart = CurrentLineStart,
                Value = date,
            };
        }

        private Token NextIdentifier()
        {
            int start = Index;
            string id = GetIdentifierName();

            Token result;

            if (id.CIEquals("true"))
                result = new TrueLiteralToken();
            else if (id.CIEquals("null"))
                result = new NullLiteralToken();
            else if (id.CIEquals("false"))
                result = new FalseLiteralToken();
            else if (id.CIEquals("empty"))
                result = new EmptyLiteralToken();
            else if (id.CIEquals("nothing"))
                result = new NothingLiteralToken();
            else if (LexerUtils.IsKeyword(id))
                result = new KeywordToken {
                    Keyword = LexerUtils.GetKeyword(id),
                    Name = id,
                };
            else if (LexerUtils.IsKeywordAsIdentifier(id))
                result = new KeywordOrIdentifierToken
                {
                    Keyword = LexerUtils.GetKeywordAsIdentifier(id),
                    Name = id,
                };
            else
                result = new IdentifierToken { Name = id };

            result.Start = start;
            result.End = Index;
            result.LineNumber = CurrentLine;
            result.LineStart = CurrentLineStart;

            return result;
        }

        private string GetIdentifierName()
        {
            int start = Index;
            while (!IsEof())
            {
                char c = Code.GetChar(Index);
                if (CharUtils.IsIdentifier(c))
                {
                    ++Index;
                }
                else
                {
                    break;
                }
            }

            return Code.Slice(start, Index);
        }

        private Token NextStringLiteral()
        {
            bool err = true;
            int start = Index++;
            var str = GetStringBuilder();

            while (!IsEof())
            {
                char c = Code.GetChar(Index);
                if (c == '"')
                {
                    c = Code.GetChar(++Index);
                    if (c == '"')
                    {
                        ++Index;
                        str.Append(c);
                    }
                    else
                    {
                        err = false;
                        break;
                    }
                }
                else if (CharUtils.IsNewLine(c))
                {
                    break;
                }
                else
                {
                    ++Index;
                    str.Append(c);
                }
            }

            if (err)
            {
                throw VBSyntaxError(VBSyntaxErrorCode.UnterminatedStringConstant);
            }

            return new StringLiteralToken
            {
                Start = start,
                End = Index - 1,
                LineNumber = CurrentLine,
                LineStart = CurrentLineStart,
                Value = str.ToString(),
            };
        }

        private Token NextNumericLiteral()
        {
            int start = Index;
            char c = Code.GetChar(Index);
            char next = Code.GetChar(Index + 1);

            string? dec = null;
            StringBuilder? fstr = null;

            if (c != '.')
            {
                if (c == '&')
                {
                    if (CharUtils.Equals(next, 'h'))
                    {
                        return NextHexIntLiteral();
                    }
                    else if (CharUtils.IsOctDigit(next))
                    {
                        return NextOctIntLiteral();
                    }
                    else
                    {
                        throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError);
                    }
                }
                else
                {
                    dec = GetDecStr();

                    if (CharUtils.IsIdentifierStart(Code.GetChar(Index)))
                    {
                        throw VBSyntaxError(VBSyntaxErrorCode.ExpectedEndOfStatement);
                    }
                }
            }

            c = Code.GetChar(Index);
            if (c == '.')
            {
                ++Index;
                fstr ??= GetStringBuilder();
                fstr.Append('.').Append(GetDecStr());
                c = Code.GetChar(Index);
            }

            if (CharUtils.Equals(c, 'e'))
            {
                fstr ??= GetStringBuilder();
                fstr.Append('e');

                c = Code.GetChar(++Index);
                if (c == '+' || c == '-')
                {
                    ++Index;
                    fstr.Append(c);
                }

                c = Code.GetChar(Index);
                if (CharUtils.IsDecDigit(c))
                {
                    fstr.Append(GetDecStr());
                }
                else
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.InvalidNumber);
                }
            }

            c = Code.GetChar(Index);
            if (CharUtils.IsIdentifierStart(c))
            {
                throw VBSyntaxError(VBSyntaxErrorCode.ExpectedEndOfStatement);
            }

            if (fstr != null && dec != null)
                fstr.Insert(0, dec);

            if (fstr != null)
            {
                return new FloatLiteralToken
                {
                    Start = start,
                    End = Index,
                    LineNumber = CurrentLine,
                    LineStart = CurrentLineStart,
                    Value = ParseDouble(fstr.ToString()),
                };
            }

            var result = ParseInteger(dec!, 10);
            result.Start = start;

            return result;
        }

        private double ParseDouble(string str)
        {
            if (!Double.TryParse(str, NumberStyles.Any,
                CultureInfo.InvariantCulture, out var result))
            {
                throw VBSyntaxError(VBSyntaxErrorCode.InvalidNumber);
            }

            return result;
        }

        private string GetStrByPredicate(Func<char, bool> predicate)
        {
            int start = Index;
            char c = Code.GetChar(Index);
            while (predicate(c))
            {
                c = Code.GetChar(++Index);
            }

            return Code.Slice(start, Index);
        }

        private string GetDecStr() => GetStrByPredicate(CharUtils.IsDecDigit);
        
        private string GetOctStr() => GetStrByPredicate(CharUtils.IsOctDigit);

        private string GetHexStr() => GetStrByPredicate(CharUtils.IsHexDigit);

        private LiteralToken ParseInteger(string str, int fromBase)
        {
            LiteralToken? result = null;

            try
            {
                int value = Convert.ToInt32(str, fromBase);
                
                result = fromBase switch
                {
                    8 => new OctIntegerLiteralToken { Value = value },
                    10 => new DecIntegerLiteralToken { Value = value },
                    16 => new HexIntegerLiteralToken { Value = value },
                    _ => throw new NotSupportedException(),
                };

                result.End = Index;
                result.LineNumber = CurrentLine;
                result.LineStart = CurrentLineStart;
            }
            catch (OverflowException)
            {

            }

            if (result == null)
            {
                if (fromBase == 8 || fromBase == 16)
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError);
                }

                result = new FloatLiteralToken
                {
                    End = Index,
                    LineNumber = CurrentLine,
                    LineStart = CurrentLineStart,
                    Value = ParseDouble(str),
                };
            }

            return result;
        }

        private LiteralToken NextOctIntLiteral()
        {
            int start = Index;
            ++Index;

            string str = GetOctStr();
            char c = Code.GetChar(Index);

            if (CharUtils.IsDecDigit(c) && !CharUtils.IsOctDigit(c))
            {
                throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError);
            }

            if (CharUtils.IsIdentifierStart(c))
            {
                throw VBSyntaxError(VBSyntaxErrorCode.ExpectedEndOfStatement);
            }

            var result = ParseInteger(str, 8);
            result.Start = start;

            return result;
        }

        private LiteralToken NextHexIntLiteral()
        {
            int start = Index;
            Index += 2;
            
            string str = GetHexStr();
            char c = Code.GetChar(Index);

            if (CharUtils.IsIdentifierStart(c))
            {
                throw VBSyntaxError(VBSyntaxErrorCode.ExpectedEndOfStatement);
            }

            var result = ParseInteger(str, 16);
            result.Start = start;

            return result;
        }

        private Token? NextComment()
        {
            while (!IsEof())
            {
                char c = Code.GetChar(Index);
                if (c == '\'')
                {
                    ++Index;
                    return NextComment(1, false);
                }
                else if (CharUtils.Equals(c, 'r'))
                {
                    char c2 = Code.GetChar(Index + 1);
                    char c3 = Code.GetChar(Index + 2);
                    char c4 = Code.GetChar(Index + 3);
                    if (CharUtils.Equals(c2, 'e') &&
                        CharUtils.Equals(c3, 'm') &&
                        CharUtils.IsWhiteSpace(c4))
                    {
                        Index += 3;
                        return NextComment(3, true);
                    }
                    else
                    {
                        break;
                    }
                }
                else
                {
                    break;
                }
            }

            return null;
        }

        private Token NextComment(int offset, bool isRem)
        {
            int start = Index - offset;
            var str = GetStringBuilder();
            while (!IsEof())
            {
                char c = Code.GetChar(Index);
                if (CharUtils.IsNewLine(c))
                {
                    break;
                }
                ++Index;
                str.Append(c);
            }

            return new CommentToken
            {
                Comment = str.ToString(),
                IsRem = isRem,
                Start = start,
                LineNumber = CurrentLine,
                LineStart = CurrentLineStart,
                End = Index,
            };
        }

        private Token NextLineTermination()
        {
            int start = Index;
            int line = CurrentLine;
            bool isColon = false;

            while (!IsEof())
            {
                char c = Code.GetChar(Index);
                if (CharUtils.IsLineTerminator(c))
                {
                    if (c == '\r' && Code.GetChar(Index + 1) == '\n')
                    {
                        ++Index;
                    }
                    
                    ++Index;
                    isColon |= (c == ':');
                    
                    if (c != ':')
                    {
                        ++CurrentLine;
                        CurrentLineStart = Index;
                    }
                }
                else
                {
                    break;
                }

                SkipWhitespaces();
            }

            Token token = isColon && line == CurrentLine
                ? new ColonLineTerminationToken()
                : new LineTerminationToken();

            token.Start = start;
            token.End = Index;
            token.LineNumber = CurrentLine - 1;
            token.LineStart = CurrentLineStart;

            return token;
        }

        private Token NextPunctuation()
        {
            int start = Index;
            char c = Code.GetChar(Index);
            char next = Code.GetChar(Index + 1);

            Punctuation? type = c switch
            {
                '(' => Punctuation.LParen,
                ')' => Punctuation.RParen,
                '.' => Punctuation.Dot,
                ',' => Punctuation.Comma,
                '+' => Punctuation.Plus,
                '-' => Punctuation.Minus,
                '/' => Punctuation.Slash,
                '\\' => Punctuation.Backslash,
                '*' => Punctuation.Star,
                '&' => Punctuation.Amp,
                '^' => Punctuation.Exp,
                _ => null,
            };

            if (type == null)
            {
                switch (c)
                {
                    case '=':
                        switch (next)
                        {
                            case '<':
                                ++Index;
                                type = Punctuation.LessOrEqual;
                                break;
                            case '>':
                                ++Index;
                                type = Punctuation.GreaterOrEqual;
                                break;
                            default:
                                type = Punctuation.Equal;
                                break;
                        }
                        break;
                    case '<':
                        switch (next)
                        {
                            case '=':
                                ++Index;
                                type = Punctuation.LessOrEqual;
                                break;
                            case '>':
                                ++Index;
                                type = Punctuation.NotEqual;
                                break;
                            default:
                                type = Punctuation.Less;
                                break;
                        }
                        break;
                    case '>':
                        switch (next)
                        {
                            case '=':
                                ++Index;
                                type = Punctuation.GreaterOrEqual;
                                break;
                            case '<':
                                ++Index;
                                type = Punctuation.NotEqual;
                                break;
                            default:
                                type = Punctuation.Greater;
                                break;
                        }
                        break;
                }
            }

            if (type == null)
            {
                throw VBSyntaxError(VBSyntaxErrorCode.InvalidCharacter);
            }

            return new PunctuationToken
            {
                Start = start,
                End = ++Index,
                LineNumber = CurrentLine,
                LineStart = CurrentLineStart,
                Type = type.Value,
            };
        }

        private void SkipNewline()
        {
            char c = Code.GetChar(Index++);
            if (CharUtils.IsNewLine(c))
            {
                if (c == '\r' && Code.GetChar(Index) == '\n')
                {
                    ++Index;
                }
                ++CurrentLine;
                CurrentLineStart = Index;
            }
        }

        private bool IsEof() => Index >= _length;

        private Exception VBSyntaxError(VBSyntaxErrorCode code)
        {
            return new VBSyntaxErrorException(
                code, CurrentLine, Index - CurrentLineStart);
        }

        private StringBuilder GetStringBuilder()
        {
            _sb.Clear();
            return _sb;
        }
    }
}
