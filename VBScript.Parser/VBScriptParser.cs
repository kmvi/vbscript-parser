using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VBScript.Parser.Ast;

namespace VBScript.Parser
{
    public class VBScriptParser
    {
        private readonly ParsingOptions _options;
        private readonly VBScriptLexer _lexer;
        private Token _next;
        private Marker _startMarker;
        private Marker _lastMarker;
        private List<Comment> _comments = new();
        private bool _inWithBlock;

        public VBScriptParser(string code)
            : this(code, new())
        {

        }

        public VBScriptParser(string code, ParsingOptions options)
        {
            if (code == null)
            {
                throw new ArgumentNullException(nameof(code));
            }

            _options = options ?? throw new ArgumentNullException(nameof(options));
            _lexer = new VBScriptLexer(code);

            _startMarker = new Marker(0, _lexer.CurrentLine, 0);
            _lastMarker = new Marker(0, _lexer.CurrentLine, 0);
            _lastMarker = new Marker(0, _lexer.CurrentLine, 0);

            _next = new InvalidToken();
        }

        public Program Parse()
        {
            void Reset()
            {
                _lexer.Reset();

                _startMarker = new Marker(0, _lexer.CurrentLine, 0);
                _lastMarker = new Marker(0, _lexer.CurrentLine, 0);

                _next = _lexer.NextToken();

                _lastMarker = new Marker(_lexer.Index, _lexer.CurrentLine, _lexer.LineIndex);
            }

            bool ParseOptionExplicit()
            {
                if (OptKeyword(Keyword.Option))
                {
                    ExpectKeyword(Keyword.Explicit);

                    SkipComments();
                    ExpectEofOrLineTermination();

                    return true;
                }

                return false;
            }

            Reset();
            
            var marker = CreateMarker();
            SkipCommentsAndNewlines();
            var program = new Program(ParseOptionExplicit());
            
            while (true)
            {
                SkipCommentsAndNewlines();
                if (MatchEof())
                {
                    break;
                }

                program.Body.Add(ParseGlobalStatement());
            }

            SkipCommentsAndNewlines();

            program.Comments.AddRange(_comments);

            return Finalize(marker, program);
        }

        #region Statements

        private Statement ParseGlobalStatement()
        {
            Statement ParseGlobalBlockStatement()
                => ParseBlockStatement(inGlobal: true);
            
            var marker = CreateMarker();

            Statement stmt;
            var token = _next;
            if (MatchKeyword(Keyword.Class))
            {
                stmt = ParseClassDeclaration();
            }
            else if (MatchKeyword(Keyword.Sub))
            {
                stmt = ParseSubDeclaration(MethodAccessModifier.None, false, false);
            }
            else if (MatchKeyword(Keyword.Function))
            {
                stmt = ParseFunctionDeclaration(MethodAccessModifier.None, false, false);
            }
            else if (MatchKeyword(Keyword.Private) || MatchKeyword(Keyword.Public))
            {
                stmt = ParsePublicOrPrivate(true, false);
            }
            else
            {
                stmt = ParseGlobalBlockStatement();
            }

            return Finalize(marker, stmt);
        }

        private ClassDeclaration ParseClassDeclaration()
        {
            static bool IsDefault(Statement memberStmt) => memberStmt switch
            {
                PropertyDeclaration p => p.AccessModifier == MethodAccessModifier.PublicDefault,
                SubDeclaration p => p.AccessModifier == MethodAccessModifier.PublicDefault,
                _ => false,
            };

            var marker = CreateMarker();

            ExpectKeyword(Keyword.Class);

            var id = ParseIdentifier();

            SkipComments();
            ExpectLineTermination();

            var stmt = new ClassDeclaration(id);

            while (true)
            {
                SkipCommentsAndNewlines();
                if (MatchEof() || MatchKeyword(Keyword.End))
                {
                    break;
                }

                Statement memberStmt;
                if (MatchKeyword(Keyword.Public) || MatchKeyword(Keyword.Private))
                {
                    memberStmt = ParsePublicOrPrivate(false, false);
                }
                else if (MatchKeyword(Keyword.Dim))
                {
                    memberStmt = ParseVariablesDeclaration();
                }
                else if (MatchKeyword(Keyword.Const))
                {
                    memberStmt = ParseConstDeclaration(MemberAccessModifier.None);
                }
                else if (MatchKeyword(Keyword.Function))
                {
                    memberStmt = ParseFunctionDeclaration(MethodAccessModifier.None, true, false);
                }
                else if (MatchKeyword(Keyword.Sub))
                {
                    memberStmt = ParseSubDeclaration(MethodAccessModifier.None, true, false);
                }
                else if (MatchKeyword(Keyword.Property))
                {
                    memberStmt = ParsePropertyDeclaration(MethodAccessModifier.None);
                }
                else if (MatchKeyword(Keyword.Default))
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.DefaultMustAlsoSpecifyPublic);
                }
                else
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError); // TODO
                }

                SkipComments();
                ExpectLineTermination();

                if (IsDefault(memberStmt) && stmt.Members.Any(IsDefault))
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.CannotHaveMultipleDefault);
                }

                if (memberStmt is SubDeclaration s)
                {
                    bool isEvent = s is InitializeSubDeclaration || s is TerminateSubDeclaration;
                    if (isEvent && s.Parameters.Count != 0)
                    {
                        throw VBSyntaxError(VBSyntaxErrorCode.ClassInitializeOrTerminateDoNotHaveArguments);
                    }
                }

                stmt.Members.Add(memberStmt);
            }

            ExpectEnd();
            ExpectKeyword(Keyword.Class, VBSyntaxErrorCode.ExpectedClass);

            return Finalize(marker, stmt);
        }

        private Statement ParsePublicOrPrivate(bool inGlobal, bool inlineOnly)
        {
            var token1 = (KeywordToken)_next;
            Move();
            var token2 = _next;

            Statement memberStmt;
            bool isDefault = OptKeyword(Keyword.Default);
            if (isDefault && token1.Keyword == Keyword.Private)
            {
                throw VBSyntaxError(VBSyntaxErrorCode.DefaultMustAlsoSpecifyPublic);
            }

            if (MatchKeyword(Keyword.Sub))
            {
                memberStmt = ParseSubDeclaration(GetMethodAccessModifier(token1, token2), !inGlobal, inlineOnly);
            }
            else if (MatchKeyword(Keyword.Function))
            {
                memberStmt = ParseFunctionDeclaration(GetMethodAccessModifier(token1, token2), !inGlobal, inlineOnly);
            }
            else if ((!inGlobal || !isDefault) && MatchKeyword(Keyword.Property))
            {
                // TODO: 'property' is valid name of a field
                memberStmt = ParsePropertyDeclaration(GetMethodAccessModifier(token1, token2));
            }
            else if (!isDefault && MatchKeyword(Keyword.Const))
            {
                memberStmt = ParseConstDeclaration(GetMemberAccessModifier(token1));
            }
            else if (!isDefault && MatchIdentifier())
            {
                memberStmt = ParseFieldsDeclaration(GetFieldAccessModifier(token1));
            }
            else if (isDefault)
            {
                throw VBSyntaxError(VBSyntaxErrorCode.ExpectedEndOfStatement);
            }
            else
            {
                throw VBSyntaxError(VBSyntaxErrorCode.ExpectedIdentifier);
            }

            return memberStmt;
        }

        private FieldsDeclaration ParseFieldsDeclaration(FieldAccessModifier modifier)
        {
            var marker = CreateMarker();

            var stmt = new FieldsDeclaration(modifier);
            stmt.Fields.Add(ParseFieldDeclaration());

            while (true)
            {
                SkipComments();
                if (MatchEof() || MatchLineTermination())
                {
                    break;
                }

                ExpectPunctuation(Punctuation.Comma, VBSyntaxErrorCode.ExpectedEndOfStatement);
                stmt.Fields.Add(ParseFieldDeclaration());
            }

            return Finalize(marker, stmt);
        }

        private PropertyDeclaration ParsePropertyDeclaration(MethodAccessModifier modifier)
        {
            PropertyDeclarationCtor GetPropertyCtor()
            {
                PropertyDeclarationCtor result;
                if (OptKeyword(Keyword.Get))
                {
                    result = (modifier, id) => new PropertyGetDeclaration(modifier, id);
                }
                else if (OptKeyword(Keyword.Set))
                {
                    result = (modifier, id) => new PropertySetDeclaration(modifier, id);
                }
                else if (OptKeyword(Keyword.Let))
                {
                    result = (modifier, id) => new PropertyLetDeclaration(modifier, id);
                }
                else
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.ExpectedLetGetSet);
                }

                return result;
            }

            var marker = CreateMarker();

            ExpectKeyword(Keyword.Property);

            var ctor = GetPropertyCtor();
            var id = ParseIdentifier();

            var stmt = ctor(modifier, id);

            if (OptPunctuation(Punctuation.LParen))
            {
                stmt.Parameters.AddRange(ParseParameterList());
                ExpectRParen();
            }

            if (stmt is not PropertyGetDeclaration)
            {
                if (modifier == MethodAccessModifier.PublicDefault)
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.DefaultCanOnlyBeOnPropertyGet);
                }
                if (stmt.Parameters.Count == 0)
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.PropertySetOrLetMustHaveArguments);
                }
            }

            SkipComments();
            ExpectLineTermination();

            while (true)
            {
                SkipCommentsAndNewlines();
                if (MatchEof() || MatchKeyword(Keyword.End))
                {
                    break;
                }

                if (MatchKeyword(Keyword.Const))
                {
                    stmt.Body.Add(ParseConstDeclaration(MemberAccessModifier.None));
                }
                else
                {
                    stmt.Body.Add(ParseBlockStatement());
                }
            }

            ExpectEnd();
            ExpectKeyword(Keyword.Property, VBSyntaxErrorCode.ExpectedProperty);

            return Finalize(marker, stmt);
        }

        private Statement ParseBlockStatement()
            => ParseBlockStatement(inGlobal: false);

        private Statement ParseBlockStatement(bool inGlobal)
        {
            var marker = CreateMarker();

            Statement stmt;
            if (_next is KeywordToken k)
            {
                stmt = k.Keyword switch
                {
                    Keyword.If => ParseIfStatement(),
                    Keyword.For => ParseForOrForEachStatement(),
                    Keyword.Do => ParseDoStatement(),
                    Keyword.Select => ParseSelectStatement(),
                    Keyword.While => ParseWhileStatement(),
                    Keyword.With => ParseWithStatement(),
                    Keyword.Loop => throw VBSyntaxError(VBSyntaxErrorCode.LoopWithoutDo),
                    Keyword.Next => throw VBSyntaxError(VBSyntaxErrorCode.UnexpectedNext),
                    _ => ParseInlineStatement(),
                };
            }
            else
            {
                stmt = ParseInlineStatement();
            }

            SkipComments();
            if (inGlobal)
            {
                ExpectEofOrLineTermination(VBSyntaxErrorCode.ExpectedEndOfStatement);
            }
            else
            {
                ExpectLineTermination(VBSyntaxErrorCode.ExpectedEndOfStatement);
            }

            return Finalize(marker, stmt);
        }

        private Statement ParseWhileStatement()
        {
            var marker = CreateMarker();

            ExpectKeyword(Keyword.While);

            var cond = ParseExpression();

            SkipComments();
            ExpectLineTermination();

            var stmt = new WhileStatement(cond);

            while (true)
            {
                SkipCommentsAndNewlines();
                if (MatchEof() || MatchKeyword(Keyword.WEnd))
                {
                    break;
                }
                
                stmt.Body.Add(ParseBlockStatement());
            }

            ExpectKeyword(Keyword.WEnd, VBSyntaxErrorCode.ExpectedWend);

            return Finalize(marker, stmt);
        }

        private WithStatement ParseWithStatement()
        {
            Statement ParseBlockStatementInsideWith()
                => ParseBlockStatement(inGlobal: false);

            var marker = CreateMarker();

            ExpectKeyword(Keyword.With);
            var expr = ParseExpression();

            SkipComments();
            ExpectLineTermination();

            _inWithBlock = true;
            var stmt = new WithStatement(expr);

            while (true)
            {
                SkipCommentsAndNewlines();
                if (MatchEof() || MatchKeyword(Keyword.End))
                {
                    break;
                }

                stmt.Body.Add(ParseBlockStatementInsideWith());
            }

            _inWithBlock = false;

            ExpectEnd();
            ExpectKeyword(Keyword.With, VBSyntaxErrorCode.ExpectedWith);

            return Finalize(marker, stmt);
        }

        private Statement ParseSelectStatement()
        {
            CaseStatement ParseCaseStatement(ref bool last)
            {
                var marker = CreateMarker();
                var stmt = new CaseStatement();
                if (!OptKeyword(Keyword.Else))
                {
                    stmt.Values.Add(ParseExpression());
                    while (OptPunctuation(Punctuation.Comma))
                    {
                        if (!MatchKeyword(Keyword.Else))
                        {
                            stmt.Values.Add(ParseExpression());
                        }
                        else
                        {
                            throw VBSyntaxError(VBSyntaxErrorCode.ExpectedExpression);
                        }
                    }
                }
                else
                {
                    last = true;
                }

                SkipComments();
                OptLineTermination();

                while (true)
                {
                    SkipCommentsAndNewlines();
                    if (MatchEof() || MatchKeyword(Keyword.Case) || MatchKeyword(Keyword.End))
                    {
                        break;
                    }

                    stmt.Body.Add(ParseBlockStatement());
                }

                return Finalize(marker, stmt);
            }

            var marker = CreateMarker();

            ExpectKeyword(Keyword.Select);
            ExpectKeyword(Keyword.Case);

            var cond = ParseExpression();

            SkipComments();
            ExpectLineTermination();

            var stmt = new SelectStatement(cond);

            bool last = false;
            while (true)
            {
                SkipCommentsAndNewlines();
                if (!OptKeyword(Keyword.Case))
                {
                    break;
                }
                else if (last)
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.ExpectedEnd);
                }

                stmt.Cases.Add(ParseCaseStatement(ref last));
            }

            SkipCommentsAndNewlines();

            ExpectEnd();
            ExpectKeyword(Keyword.Select, VBSyntaxErrorCode.ExpectedSelect);

            return Finalize(marker, stmt);
        }

        private Statement ParseDoStatement()
        {
            var marker = CreateMarker();

            ExpectKeyword(Keyword.Do);

            Expression? cond = null;
            var loopType = LoopType.None;
            var testType = ConditionTestType.None;
            if (OptKeyword(Keyword.While))
            {
                testType = ConditionTestType.PreTest;
                loopType = LoopType.While;
                cond = ParseExpression();
            }
            else if (OptKeyword(Keyword.Until))
            {
                testType = ConditionTestType.PreTest;
                loopType = LoopType.Until;
                cond = ParseExpression();
            }

            SkipComments();
            ExpectLineTermination(loopType == LoopType.None
                ? VBSyntaxErrorCode.ExpectedWhileUntilOrEndOfStatement
                : VBSyntaxErrorCode.SyntaxError);

            List<Statement> body = new();
            while (true)
            {
                SkipCommentsAndNewlines();
                if (MatchEof() || MatchKeyword(Keyword.Loop))
                {
                    break;
                }
                
                body.Add(ParseBlockStatement());
            }

            ExpectKeyword(Keyword.Loop, VBSyntaxErrorCode.ExpectedLoop);

            if (testType == ConditionTestType.None)
            {
                if (OptKeyword(Keyword.While))
                {
                    loopType = LoopType.While;
                    cond = ParseExpression();
                    testType = ConditionTestType.PostTest;
                }
                else if (OptKeyword(Keyword.Until))
                {
                    loopType = LoopType.Until;
                    cond = ParseExpression();
                    testType = ConditionTestType.PostTest;
                }
            }

            var stmt = new DoStatement(loopType, testType, cond!);
            stmt.Body.AddRange(body);

            return Finalize(marker, stmt);
        }

        private Statement ParseForOrForEachStatement()
        {
            var marker = CreateMarker();

            ExpectKeyword(Keyword.For);

            Statement stmt = OptKeyword(Keyword.Each)
                ? ParseForEachStatement()
                : ParseForStatement();

            return Finalize(marker, stmt);
        }

        private Statement ParseForStatement()
        {
            var marker = CreateMarker();

            var id = ParseIdentifier();
            ExpectPunctuation(Punctuation.Equal, VBSyntaxErrorCode.ExpectedEqual);
            var from = ParseExpression();
            ExpectKeyword(Keyword.To, VBSyntaxErrorCode.ExpectedTo);
            var to = ParseExpression();

            Expression? step = null;
            if (OptKeyword(Keyword.Step))
            {
                step = ParseExpression();
            }

            SkipComments();
            ExpectLineTermination();

            var stmt = new ForStatement(id, from, to, step);
            while (true)
            {
                SkipCommentsAndNewlines();
                if (MatchEof() || MatchKeyword(Keyword.Next))
                {
                    break;
                }
                stmt.Body.Add(ParseBlockStatement());
            }

            ExpectKeyword(Keyword.Next, VBSyntaxErrorCode.ExpectedNext);

            return Finalize(marker, stmt);
        }

        private ForEachStatement ParseForEachStatement()
        {
            var marker = CreateMarker();

            var id = ParseIdentifier();
            ExpectKeyword(Keyword.In, VBSyntaxErrorCode.ExpectedIn);
            var @in = ParseExpression();

            SkipComments();
            ExpectLineTermination();

            var stmt = new ForEachStatement(id, @in);
            while (true)
            {
                SkipCommentsAndNewlines();
                if (MatchEof() || MatchKeyword(Keyword.Next))
                {
                    break;
                }

                stmt.Body.Add(ParseBlockStatement());
            }

            ExpectKeyword(Keyword.Next, VBSyntaxErrorCode.ExpectedNext);

            return Finalize(marker, stmt);
        }

        private Statement ParseMultiInlineStatement(bool matchElse, int line)
        {
            bool canContinue;
            var stmts = new StatementList();

            do
            {
                stmts.Add(ParseInlineStatement());
                SkipComments();
                canContinue = OptColonLineTermination() &&
                    line == _startMarker.Line &&
                    (!MatchKeyword(Keyword.Else) || !matchElse) &&
                    !MatchKeyword(Keyword.End);
            }
            while (canContinue && !MatchEof());

            return stmts.Count == 1 ? stmts[0] : stmts;
        }

        private IfStatement ParseIfStatement(bool inlineOnly = false)
        {
            var marker = CreateMarker();
            int line = _startMarker.Line;

            ExpectKeyword(Keyword.If);
            var test = ParseExpression();
            ExpectKeyword(Keyword.Then, VBSyntaxErrorCode.ExpectedThen);

            SkipComments();
            bool inline = !OptLineTermination() || line == _startMarker.Line;

            if (inlineOnly && !inline)
            {
                throw VBSyntaxError(VBSyntaxErrorCode.MustBeFirstStatementOnTheLine);
            }

            Statement cons;
            Statement? alt = null;

            if (!inline)
            {
                var block = new StatementList();
                while (true)
                {
                    SkipCommentsAndNewlines();
                    if (MatchEof() || MatchKeyword(Keyword.End) ||
                        MatchKeyword(Keyword.Else) || MatchKeyword(Keyword.ElseIf))
                    {
                        break;
                    }
                    block.Add(ParseBlockStatement());
                }

                if (MatchKeyword(Keyword.Else))
                {
                    alt = ParseElseStatement();
                }
                else if (MatchKeyword(Keyword.ElseIf))
                {
                    block.Add(ParseElseIfStatement());
                }

                ExpectEnd();
                ExpectKeyword(Keyword.If, VBSyntaxErrorCode.ExpectedIf);
                
                cons = block;
            }
            else
            {
                cons = ParseMultiInlineStatement(true, line);

                if (OptKeyword(Keyword.Else))
                {
                    alt = ParseMultiInlineStatement(false, line);
                }
                if (OptKeyword(Keyword.End))
                {
                    ExpectKeyword(Keyword.If, VBSyntaxErrorCode.ExpectedIf);
                }
            }

            return Finalize(marker, new IfStatement(test, cons, alt));
        }

        private Statement ParseElseIfStatement()
        {
            var marker = CreateMarker();
            var line = _startMarker.Line;

            ExpectKeyword(Keyword.ElseIf);

            var test = ParseExpression();
            ExpectKeyword(Keyword.Then, VBSyntaxErrorCode.ExpectedThen);

            SkipComments();
            bool inline = !OptLineTermination();

            Statement cons;
            Statement? alt = null;
            if (!inline)
            {
                SkipCommentsAndNewlines();

                var block = new StatementList();
                while (true)
                {
                    SkipCommentsAndNewlines();
                    if (MatchEof() || MatchKeyword(Keyword.End) ||
                        MatchKeyword(Keyword.Else) || MatchKeyword(Keyword.ElseIf))
                    {
                        break;
                    }

                    block.Add(ParseBlockStatement());
                }

                if (MatchKeyword(Keyword.Else))
                {
                    alt = ParseElseStatement();
                }
                else if (MatchKeyword(Keyword.ElseIf))
                {
                    block.Add(ParseElseIfStatement());
                }

                cons = block;
            }
            else
            {
                cons = ParseMultiInlineStatement(true, line);

                SkipCommentsAndNewlines();
                if (MatchKeyword(Keyword.Else))
                {
                    alt = ParseElseStatement();
                }
                else if (MatchKeyword(Keyword.ElseIf))
                {
                    var block = new StatementList();
                    block.Add(cons);
                    block.Add(ParseElseIfStatement());
                    cons = block;
                }
            }

            return Finalize(marker, new ElseIfStatement(test, cons, alt));
        }

        private Statement ParseElseStatement()
        {
            var marker = CreateMarker();
            int line = _startMarker.Line;

            ExpectKeyword(Keyword.Else);

            SkipComments();
            bool inline = !OptLineTermination();
            
            Statement stmt;
            if (!inline)
            {
                var elseBlock = new StatementList();
                while (true)
                {
                    SkipCommentsAndNewlines();
                    if (MatchEof() || MatchKeyword(Keyword.End))
                    {
                        break;
                    }
                    elseBlock.Add(ParseBlockStatement());
                }
                stmt = elseBlock;
            }
            else
            {
                stmt = ParseMultiInlineStatement(false, line);
            }

            return Finalize(marker, stmt);
        }

        private ReDimStatement ParseReDimStatement()
        {
            ReDimDeclaration ParseReDimDeclaration()
            {
                var marker = CreateMarker();

                var id = ParseIdentifier();
                var redim = new ReDimDeclaration(id);
                
                ExpectPunctuation(Punctuation.LParen, VBSyntaxErrorCode.ExpectedLParen);
                redim.ArrayDims.Add(ParseExpression());
                while (OptPunctuation(Punctuation.Comma))
                {
                    redim.ArrayDims.Add(ParseExpression());
                }

                ExpectRParen();
                
                return Finalize(marker, redim);
            }

            var marker = CreateMarker();

            ExpectKeyword(Keyword.ReDim);

            bool preserve = OptKeyword(Keyword.Preserve);
            var result = new ReDimStatement(preserve);
            result.ReDims.Add(ParseReDimDeclaration());
            while (OptPunctuation(Punctuation.Comma))
            {
                result.ReDims.Add(ParseReDimDeclaration());
            }

            return Finalize(marker, result);
        }

        private Statement ParseInlineStatement()
        {
            var marker = CreateMarker();

            Statement stmt;
            if (_next is KeywordToken k)
            {
                stmt = k.Keyword switch
                {
                    Keyword.Dim => ParseVariablesDeclaration(),
                    Keyword.ReDim => ParseReDimStatement(),
                    Keyword.Const => ParseConstDeclaration(MemberAccessModifier.None),
                    Keyword.On => ParseOnErrorStatement(),
                    Keyword.Exit => ParseExitStatement(),
                    Keyword.Erase => ParseEraseStatement(),
                    Keyword.Set => ParseSetAssignmentStatement(),
                    Keyword.Call => ParseCallStatement(),
                    Keyword.If => ParseIfStatement(inlineOnly: true),
                    Keyword.ElseIf => throw VBSyntaxError(VBSyntaxErrorCode.MustBeFirstStatementOnTheLine),
                    Keyword.Public => ParsePublicOrPrivate(inGlobal: false, inlineOnly: true),
                    Keyword.Private => ParsePublicOrPrivate(inGlobal: false, inlineOnly: true),
                    Keyword.Sub => ParseSubDeclaration(MethodAccessModifier.None, isMethod: false, inlineOnly: true),
                    Keyword.Function => ParseFunctionDeclaration(MethodAccessModifier.None, isMethod:  false, inlineOnly: true),
                    _ => throw VBSyntaxError(VBSyntaxErrorCode.ExpectedStatement),
                };
            }
            else if (_inWithBlock && MatchPunctuation(Punctuation.Dot))
            {
                stmt = ParseAssignmentOrCallStatement();
            }
            else if (MatchIdentifier())
            {
                stmt = ParseAssignmentOrCallStatement();
            }
            else
            {
                throw VBSyntaxError(VBSyntaxErrorCode.ExpectedStatement);
            }

            return Finalize(marker, stmt);
        }

        private Statement ParseCallStatement()
        {
            var marker = CreateMarker();

            ExpectKeyword(Keyword.Call);

            return Finalize(marker, new CallStatement(ParseLeftExpression()));
        }

        private Statement ParseAssignmentOrCallStatement()
        {
            Statement ParseArgs(Expression callee, Expression firstArg)
            {
                Statement stmt;
                var callstmt = new CallSubStatement(callee);
                callstmt.Arguments.Add(firstArg);
                while (OptPunctuation(Punctuation.Comma))
                {
                    bool isEmptyValue = MatchPunctuation(Punctuation.Comma);
                    var arg = isEmptyValue ? ParseMissingValueExpression() : ParseExpression();
                    callstmt.Arguments.Add(arg);
                }
                stmt = callstmt;
                return stmt;
            }

            var left = ParseLeftExpression();

            Statement stmt;
            if (OptPunctuation(Punctuation.Equal))
            {
                CheckNotMeIdentifier(left);
                Expression right = ParseExpression();
                stmt = new AssignmentStatement(left, right, false);
            }
            else if (MatchLineTermination())
            {
                if (left is IndexOrCallExpression expr && expr.Indexes.Count <= 1)
                {
                    var callstmt = new CallSubStatement(expr.Object);
                    if (expr.Indexes.Count != 0)
                    {
                        callstmt.Arguments.Add(expr.Indexes[0]);
                    }
                    stmt = callstmt;
                }
                else
                {
                    stmt = new CallSubStatement(left);
                }
            }
            else if (MatchPunctuation(Punctuation.Comma))
            {
                if (left is IndexOrCallExpression expr && expr.Indexes.Count <= 1)
                {
                    if (expr.Indexes.Count == 0)
                    {
                        throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError); // TODO:
                    }

                    stmt = ParseArgs(expr.Object, expr.Indexes[0]);
                }
                else if (left is Identifier || left is MemberExpression)
                {
                    stmt = ParseArgs(left, ParseMissingValueExpression());
                }
                else
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.ExpectedEndOfStatement);
                }
            }
            else
            {
                var callstmt = new CallSubStatement(left);

                SkipComments();
                if (!MatchLineTermination() && !MatchEof())
                {
                    callstmt.Arguments.Add(ParseExpression());

                    while (OptPunctuation(Punctuation.Comma))
                    {
                        bool isEmptyValue = MatchPunctuation(Punctuation.Comma);
                        var expr = isEmptyValue ? ParseMissingValueExpression() : ParseExpression();
                        callstmt.Arguments.Add(expr);
                    }
                }

                stmt = callstmt;
            }

            return stmt;
        }

        private ConstsDeclaration ParseConstDeclaration(MemberAccessModifier modifier)
        {
            ConstDeclaration AddConstDeclarator()
            {
                var marker = CreateMarker();
                var id = ParseIdentifier();
                ExpectPunctuation(Punctuation.Equal, VBSyntaxErrorCode.ExpectedEqual);
                var expr = ParseConstInitExpression();
                return Finalize(marker, new ConstDeclaration(id, expr));
            }

            var marker = CreateMarker();

            ExpectKeyword(Keyword.Const);

            var result = new ConstsDeclaration(modifier);
            result.Declarations.Add(AddConstDeclarator());

            while (true)
            {
                SkipComments();
                if (MatchEof() || MatchLineTermination())
                {
                    break;
                }

                ExpectPunctuation(Punctuation.Comma, VBSyntaxErrorCode.ExpectedEndOfStatement);
                result.Declarations.Add(AddConstDeclarator());
            }

            return Finalize(marker, result);
        }

        private Statement ParseEraseStatement()
        {
            var marker = CreateMarker();
            ExpectKeyword(Keyword.Erase);
            var id = ParseIdentifier();
            return Finalize(marker, new EraseStatement(id));
        }

        private Statement ParseExitStatement()
        {
            var marker = CreateMarker();

            ExpectKeyword(Keyword.Exit);

            Statement stmt;
            if (OptKeyword(Keyword.Do))
            {
                stmt = new ExitDoStatement();
            }
            else if (OptKeyword(Keyword.For))
            {
                stmt = new ExitForStatement();
            }
            else if (OptKeyword(Keyword.Sub))
            {
                stmt = new ExitSubStatement();
            }
            else if (OptKeyword(Keyword.Function))
            {
                stmt = new ExitFunctionStatement();
            }
            else if (OptKeyword(Keyword.Property))
            {
                stmt = new ExitPropertyStatement();
            }
            else
            {
                throw VBSyntaxError(VBSyntaxErrorCode.InvalidExitStatement);
            }

            return Finalize(marker, stmt);
        }

        private FunctionDeclaration ParseFunctionDeclaration(MethodAccessModifier modifier, bool isMethod, bool inlineOnly)
        {
            static FunctionDeclaration Ctor(MethodAccessModifier modifier, Identifier id, Statement body)
                => new FunctionDeclaration(modifier, id, body);
            
            return ParseProcedure(Ctor, Keyword.Function, modifier, isMethod, inlineOnly);
        }

        private SubDeclaration ParseSubDeclaration(MethodAccessModifier modifier, bool isMethod, bool inlineOnly)
        {
            static SubDeclaration Ctor(MethodAccessModifier modifier, Identifier id, Statement body)
            {
                if (id.Name.CIEquals(InitializeSubDeclaration.Name))
                {
                    return new InitializeSubDeclaration(modifier, body);
                }
                else if (id.Name.CIEquals(TerminateSubDeclaration.Name))
                {
                    return new TerminateSubDeclaration(modifier, body);
                }
                else
                {
                    return new SubDeclaration(modifier, id, body);
                }
            }

            return ParseProcedure(Ctor, Keyword.Sub, modifier, isMethod, inlineOnly);
        }

        private T ParseProcedure<T>(
            SubDeclarationCtor<T> ctor, Keyword kw,
            MethodAccessModifier modifier, bool isMethod, bool inlineOnly)
            where T : ProcedureDeclaration
        {
            var marker = CreateMarker();
            int line = _startMarker.Line;

            ExpectKeyword(kw);

            var id = ParseIdentifier();

            bool hasLParen = false;
            List<Parameter> args = new();
            if (OptPunctuation(Punctuation.LParen))
            {
                hasLParen = true;
                args = ParseParameterList();
                ExpectRParen();
            }

            SkipComments();
            bool inline = MatchColonLineTermination() || !MatchLineTermination();
            hasLParen |= MatchColonLineTermination();
            OptLineTermination();

            if (!inline && inlineOnly)
            {
                throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError);
            }

            if (inline && !hasLParen)
            {
                throw VBSyntaxError(VBSyntaxErrorCode.ExpectedLParen);
            }

            Statement body;
            if (inline)
            {
                body = ParseMultiInlineStatement(false, line);
            }
            else
            {
                var list = new StatementList();
                while (true)
                {
                    SkipCommentsAndNewlines();
                    if (MatchEof() || MatchKeyword(Keyword.End))
                    {
                        break;
                    }

                    list.Add(ParseBlockStatement());
                }
                body = list;
            }

            var stmt = ctor(modifier, id, body);
            stmt.Parameters.AddRange(args);

            ExpectEnd();
            
            var code = kw == Keyword.Sub ? VBSyntaxErrorCode.ExpectedSub : VBSyntaxErrorCode.ExpectedFunction;
            ExpectKeyword(kw, code);

            return Finalize(marker, stmt);
        }

        private List<Parameter> ParseParameterList()
        {
            var result = new List<Parameter>();

            if (MatchKeyword(Keyword.ByRef) || MatchKeyword(Keyword.ByVal) || MatchIdentifier())
            {
                result.Add(ParseParameter());
            }

            if (result.Count != 0)
            {
                while (OptPunctuation(Punctuation.Comma))
                {
                    result.Add(ParseParameter());
                }
            }

            return result;
        }

        private Parameter ParseParameter()
        {
            var marker = CreateMarker();

            var modifier = ParameterModifier.None;
            if (OptKeyword(Keyword.ByRef))
            {
                modifier = ParameterModifier.ByRef;
            }
            else if (OptKeyword(Keyword.ByVal))
            {
                modifier = ParameterModifier.ByVal;
            }

            var id = ParseIdentifier(); // todo: code

            bool parens = false;
            if (OptPunctuation(Punctuation.LParen))
            {
                ExpectRParen();
                parens = true;
            }

            return Finalize(marker, new Parameter(id, modifier, parens));
        }

        private static MethodAccessModifier GetMethodAccessModifier(Token token1, Token token2)
        {
            return MatchKeyword(token1) switch
            {
                Keyword.Private => MethodAccessModifier.Private,
                Keyword.Public => MatchKeyword(token2) switch
                {
                    Keyword.Default => MethodAccessModifier.PublicDefault,
                    _ => MethodAccessModifier.Public,
                },
                _ => MethodAccessModifier.None,
            };
        }

        private static MemberAccessModifier GetMemberAccessModifier(Token token)
        {
            return MatchKeyword(token) switch
            {
                Keyword.Private => MemberAccessModifier.Private,
                Keyword.Public => MemberAccessModifier.Public,
                _ => MemberAccessModifier.None,
            };
        }

        private FieldAccessModifier GetFieldAccessModifier(Token token)
        {
            return MatchKeyword(token) switch {
                Keyword.Private => FieldAccessModifier.Private,
                Keyword.Public => FieldAccessModifier.Public,
                _ => throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError), // TODO
            };
        }

        private Statement ParseOnErrorStatement()
        {
            var marker = CreateMarker();

            ExpectKeyword(Keyword.On);
            ExpectKeyword(Keyword.Error);

            Statement stmt;
            if (OptKeyword(Keyword.Resume))
            {
                ExpectKeyword(Keyword.Next);
                stmt = new OnErrorResumeNextStatement();
            }
            else
            {
                ExpectKeyword(Keyword.Goto);
                int i = ExpectInteger(VBSyntaxErrorCode.SyntaxError);
                if (i != 0)
                {
                    throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError);
                }
                
                stmt = new OnErrorGoTo0Statement();
            }

            return Finalize(marker, stmt);
        }

        private AssignmentStatement ParseSetAssignmentStatement()
        {
            var marker = CreateMarker();

            bool set = OptKeyword(Keyword.Set);
            Expression left = ParseLeftExpression();
            CheckNotMeIdentifier(left);
            ExpectPunctuation(Punctuation.Equal, VBSyntaxErrorCode.ExpectedEqual);
            Expression right = ParseExpression();

            return Finalize(marker, new AssignmentStatement(left, right, set));
        }

        private void CheckNotMeIdentifier(Expression left)
        {
            if (left is Identifier id && id.Name.CIEquals("me"))
            {
                throw VBSyntaxError(VBSyntaxErrorCode.InvalidUseOfMeKeyword);
            }
        }

        private VariablesDeclaration ParseVariablesDeclaration()
        {
            var marker = CreateMarker();

            ExpectKeyword(Keyword.Dim);

            var stmt = new VariablesDeclaration();
            stmt.Variables.Add(ParseVariableDeclaration());

            while (true)
            {
                SkipComments();
                if (MatchEof() || MatchLineTermination())
                {
                    break;
                }

                ExpectPunctuation(Punctuation.Comma, VBSyntaxErrorCode.ExpectedEndOfStatement);
                stmt.Variables.Add(ParseVariableDeclaration());
            }

            return Finalize(marker, stmt);
        }

        private FieldDeclaration ParseFieldDeclaration()
        {
            static FieldDeclaration Ctor(Identifier id, bool isDynamicArray)
                => new FieldDeclaration(id, isDynamicArray);

            Identifier GetFieldId()
            {
                // TODO: field named 'property' is valid
                var marker = CreateMarker();
                string name = ExpectIdentifier();
                return Finalize(marker, CreateIdentifier(name));
            }

            return ParseVariableDeclarationNode(Ctor, GetFieldId);
        }

        private VariableDeclaration ParseVariableDeclaration()
        {
            static VariableDeclaration Ctor(Identifier id, bool isDynamicArray)
                => new VariableDeclaration(id, isDynamicArray);

            return ParseVariableDeclarationNode(Ctor, ParseIdentifier);
        }

        private T ParseVariableDeclarationNode<T>(
            VariableDeclarationCtor<T> ctor, Func<Identifier> identifierGetter)
            where T : VariableDeclarationNode
        {
            var marker = CreateMarker();

            var id = identifierGetter();
            var decl = ctor(id, false);

            if (OptPunctuation(Punctuation.LParen))
            {
                if (!MatchPunctuation(Punctuation.RParen))
                {
                    decl.ArrayDims.Add(ExpectInteger());

                    while (!MatchEof() && !MatchPunctuation(Punctuation.RParen))
                    {
                        ExpectPunctuation(Punctuation.Comma, VBSyntaxErrorCode.ExpectedRParen);
                        decl.ArrayDims.Add(ExpectInteger());
                    }
                }
                else
                {
                    decl = ctor(id, true);
                }

                ExpectRParen();
            }

            return Finalize(marker, decl!);
        }

        #endregion

        #region Expressions

        private Expression ParseConstInitExpression()
        {
            var marker = CreateMarker();

            Expression expr;
            if (OptPunctuation(Punctuation.Plus))
            {
                expr = new UnaryExpression(UnaryOperation.Plus, ParseConstInitExpression());
            }
            else if (OptPunctuation(Punctuation.Minus))
            {
                expr = new UnaryExpression(UnaryOperation.Minus, ParseConstInitExpression());
            }
            else if (OptPunctuation(Punctuation.LParen))
            {
                expr = ParseConstExpression();
                ExpectRParen();
            }
            else
            {
                expr = ParseConstExpression();
            }

            return Finalize(marker, expr);
        }

        private Expression ParseLeftExpression()
        {
            Identifier ParsePropertyId()
            {
                var marker = CreateMarker();
                var name = MatchIdentifier() ? ExpectIdentifier() : ExpectAnyKeywordAsIdentifier();
                return Finalize(marker, CreateIdentifier(name));
            }
            
            var marker = CreateMarker();

            Expression expr;
            if (_inWithBlock && OptPunctuation(Punctuation.Dot))
            {
                expr = new WithMemberAccessExpression(ParsePropertyId());
            }
            else
            {
                expr = ParseIdentifier();
            }

            while (true)
            {
                if (OptPunctuation(Punctuation.Dot))
                {
                    expr = new MemberExpression(expr, ParsePropertyId());
                }
                else if (OptPunctuation(Punctuation.LParen))
                {
                    var ix = new IndexOrCallExpression(expr);
                    expr = ix;

                    if (OptPunctuation(Punctuation.Comma))
                    {
                        ix.Indexes.Add(ParseMissingValueExpression());
                    }
                    else if (!MatchPunctuation(Punctuation.RParen))
                    {
                        ix.Indexes.Add(ParseExpression());
                    }

                    while (OptPunctuation(Punctuation.Comma))
                    {
                        if (MatchPunctuation(Punctuation.Comma) ||
                            MatchPunctuation(Punctuation.RParen))
                        {
                            ix.Indexes.Add(ParseMissingValueExpression());
                        }
                        else
                        {
                            ix.Indexes.Add(ParseExpression());
                        }
                    }

                    ExpectRParen();
                }
                else
                {
                    break;
                }
            }

            return Finalize(marker, expr);
        }

        private Expression ParseUnaryExpression()
        {
            var marker = CreateMarker();

            Expression expr;
            if (OptPunctuation(Punctuation.Minus))
            {
                expr = new UnaryExpression(UnaryOperation.Minus, ParseUnaryExpression());
            }
            else if (OptPunctuation(Punctuation.Plus))
            {
                expr = new UnaryExpression(UnaryOperation.Plus, ParseUnaryExpression());
            }
            else if (OptKeyword(Keyword.Not))
            {
                expr = new UnaryExpression(UnaryOperation.Not, ParseUnaryExpression());
            }
            else
            {
                expr = ParseValueExpression();
            }

            return Finalize(marker, expr);
        }

        private Expression ParseExpExpression()
        {
            var marker = CreateMarker();

            var expr = ParseUnaryExpression();
            if (OptPunctuation(Punctuation.Exp))
            {
                var right = ParseExpExpression();
                expr = new BinaryExpression(BinaryOperation.Exponentiation, expr, right);
            }

            return Finalize(marker, expr);
        }

        private Expression ParseValueExpression()
        {
            var marker = CreateMarker();

            Expression expr;
            if (_next is LiteralToken)
            {
                expr = ParseConstExpression();
            }
            else if (OptPunctuation(Punctuation.LParen))
            {
                expr = ParseExpression();
                ExpectRParen();
            }
            else if (OptKeyword(Keyword.New))
            {
                expr = new NewExpression(ParseLeftExpression());
            }
            else
            {
                expr = ParseLeftExpression();
            }

            return Finalize(marker, expr);
        }

        private Expression ParseConstExpression()
        {
            var marker = CreateMarker();

            Expression expr = _next switch
            {
                DecIntegerLiteralToken t => new IntegerLiteral(t.Value),
                StringLiteralToken t => new StringLiteral(t.Value),
                FloatLiteralToken t => new FloatLiteral(t.Value),
                DateLiteralToken t => new DateLiteral(t.Value),
                EmptyLiteralToken => new EmptyLiteral(),
                NothingLiteralToken => new NothingLiteral(),
                NullLiteralToken => new NullLiteral(),
                TrueLiteralToken => new BooleanLiteral(true),
                FalseLiteralToken => new BooleanLiteral(false),
                _ => throw VBSyntaxError(VBSyntaxErrorCode.ExpectedLiteral),
            };

            Move();

            return Finalize(marker, expr);
        }

        private static int BinaryPrecedence(Token token)
        {
            int prec = 0;
            if (token is PunctuationToken p)
            {
                switch (p.Type)
                {
                    case Punctuation.Star:
                    case Punctuation.Slash:
                        prec = 50;
                        break;
                    case Punctuation.Backslash:
                        prec = 49;
                        break;
                    case Punctuation.Plus:
                    case Punctuation.Minus:
                        prec = 47;
                        break;
                    case Punctuation.Amp:
                        prec = 46;
                        break;
                    case Punctuation.Equal:
                        prec = 30;
                        break;
                    case Punctuation.NotEqual:
                        prec = 29;
                        break;
                    case Punctuation.Less:
                        prec = 28;
                        break;
                    case Punctuation.Greater:
                        prec = 29;
                        break;
                    case Punctuation.LessOrEqual:
                        prec = 27;
                        break;
                    case Punctuation.GreaterOrEqual:
                        prec = 26;
                        break;
                }
            }

            if (token is KeywordToken k)
            {
                switch (k.Keyword)
                {
                    case Keyword.Mod:
                        prec = 48;
                        break;
                    case Keyword.Is:
                        prec = 25;
                        break;
                    case Keyword.And:
                        prec = 10;
                        break;
                    case Keyword.Or:
                        prec = 9;
                        break;
                    case Keyword.Xor:
                        prec = 8;
                        break;
                    case Keyword.Eqv:
                        prec = 7;
                        break;
                    case Keyword.Imp:
                        prec = 6;
                        break;
                }
            }

            return prec;
        }

        private Expression ParseExpression()
        {
            var marker = CreateMarker();

            return Finalize(marker, ParseBinaryExpression());
        }

        private Expression ParseBinaryExpression()
        {
            var startToken = _next;

            var expr = ParseExpExpression();

            var op = _next;
            var prec = BinaryPrecedence(op);
            if (prec > 0)
            {
                Move();

                var tokens = new Stack<Token>(new[] { startToken, _next });

                var right = ParseExpExpression();

                var stack = new Stack<object>(new object[] { expr, op, right });
                var precedences = new Stack<int>();
                precedences.Push(prec);

                while (true)
                {
                    prec = BinaryPrecedence(_next);
                    if (prec <= 0)
                    {
                        break;
                    }

                    while (stack.Count > 2 && prec <= precedences.Peek())
                    {
                        right = (Expression)stack.Pop();
                        var oper = GetBinaryOperation((Token)stack.Pop());
                        precedences.Pop();
                        expr = (Expression)stack.Pop();
                        tokens.Pop();
                        var tk = tokens.Peek();
                        var marker = CreateMarker(tk, tk.LineStart);
                        stack.Push(Finalize(marker, new BinaryExpression(oper, expr, right)));
                    }

                    stack.Push(Move());
                    precedences.Push(prec);
                    tokens.Push(_next);
                    stack.Push(ParseExpExpression());
                }

                expr = (Expression)stack.Pop();
                var lastToken = tokens.Pop();
                while (stack.Count > 1)
                {
                    var tk = tokens.Pop();
                    var lastLineStart = lastToken?.LineStart ?? 0;
                    var marker = CreateMarker(tk, lastLineStart);
                    var oper = GetBinaryOperation((Token)stack.Pop());
                    expr = Finalize(marker, new BinaryExpression(oper, (Expression)stack.Pop(), expr));
                    lastToken = tk;
                }
            }

            return expr;
        }

        private BinaryOperation GetBinaryOperation(Token operation)
        {
            if (operation is PunctuationToken p)
            {
                return p.Type switch
                {
                    Punctuation.Exp => BinaryOperation.Exponentiation,
                    Punctuation.Star => BinaryOperation.Multiplication,
                    Punctuation.Slash => BinaryOperation.Division,
                    Punctuation.Backslash => BinaryOperation.IntDivision,
                    Punctuation.Plus => BinaryOperation.Addition,
                    Punctuation.Minus => BinaryOperation.Subtraction,
                    Punctuation.Amp => BinaryOperation.Concatenation,
                    Punctuation.Equal => BinaryOperation.Equal,
                    Punctuation.NotEqual => BinaryOperation.NotEqual,
                    Punctuation.Less => BinaryOperation.Less,
                    Punctuation.Greater => BinaryOperation.Greater,
                    Punctuation.LessOrEqual => BinaryOperation.LessOrEqual,
                    Punctuation.GreaterOrEqual => BinaryOperation.GreaterOrEqual,

                    _ => throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError), // TODO: code
                };
            }

            if (operation is KeywordToken k)
            {
                return k.Keyword switch
                {
                    Keyword.Mod => BinaryOperation.Mod,
                    Keyword.Is => BinaryOperation.Is,
                    Keyword.And => BinaryOperation.And,
                    Keyword.Or => BinaryOperation.Or,
                    Keyword.Xor => BinaryOperation.Xor,
                    Keyword.Eqv => BinaryOperation.Eqv,
                    Keyword.Imp => BinaryOperation.Imp,

                    _ => throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError), // TODO: code
                };
            }

            throw VBSyntaxError(VBSyntaxErrorCode.SyntaxError); // TODO: code
        }

        private MissingValueExpression ParseMissingValueExpression()
        {
            var marker = CreateMarker();
            return Finalize(marker, new MissingValueExpression());
        }

        private Identifier ParseIdentifier()
        {
            var marker = CreateMarker();
            return Finalize(marker, CreateIdentifier(ExpectIdentifier()));
        }

        #endregion

        private bool MatchEof()
            => _next is EofToken;

        private bool MatchLineTermination()
            => _next is LineTerminationToken;

        private bool MatchColonLineTermination()
            => _next is ColonLineTerminationToken;

        private bool OptLineTermination()
        {
            if (MatchLineTermination())
            {
                Move();
                return true;
            }
            return false;
        }

        private bool OptColonLineTermination()
        {
            if (MatchColonLineTermination())
            {
                Move();
                return true;
            }
            return false;
        }

        private bool MatchKeyword(Keyword keyword)
            => _next is KeywordToken k && k.Keyword == keyword ||
               _next is KeywordOrIdentifierToken k2 && k2.Keyword == keyword;

        private static Keyword? MatchKeyword(Token token)
            => token switch
            {
                KeywordToken k => k.Keyword,
                KeywordOrIdentifierToken k => k.Keyword,
                _ => null,
            };

        private bool OptKeyword(Keyword keyword)
        {
            if (MatchKeyword(keyword))
            {
                Move();
                return true;
            }
            return false;
        }

        private bool MatchPunctuation(Punctuation type)
            => _next is PunctuationToken p && p.Type == type;

        private bool OptPunctuation(Punctuation type)
        {
            if (MatchPunctuation(type))
            {
                Move();
                return true;
            }
            return false;
        }

        private bool MatchIdentifier()
            => _next is IdentifierToken || _next is KeywordOrIdentifierToken;

        private bool MatchAnyKeywordAsIdentifier()
            => _next is IdentifierToken || _next is KeywordOrIdentifierToken;

        private string ExpectAnyKeywordAsIdentifier(VBSyntaxErrorCode code = VBSyntaxErrorCode.SyntaxError)
        {
            var token = Move();
            if (token is KeywordToken)
            {
                return token.ToString();
            }

            throw VBSyntaxError(code);
        }

        private void ExpectKeyword(Keyword keyword, VBSyntaxErrorCode code = VBSyntaxErrorCode.SyntaxError)
        {
            var token = Move();
            if (!(token is KeywordToken k && k.Keyword == keyword) &&
                !(token is KeywordOrIdentifierToken k2 && k2.Keyword == keyword))
            {
                throw VBSyntaxError(code);
            }
        }

        private void ExpectEnd(VBSyntaxErrorCode code = VBSyntaxErrorCode.ExpectedEnd)
            => ExpectKeyword(Keyword.End, code);

        private void ExpectLineTermination(VBSyntaxErrorCode code = VBSyntaxErrorCode.SyntaxError)
        {
            var token = Move();
            if (token is not LineTerminationToken)
            {
                throw VBSyntaxError(code);
            }
        }

        private void ExpectEofOrLineTermination(VBSyntaxErrorCode code = VBSyntaxErrorCode.SyntaxError)
        {
            var token = Move();
            if (token is not LineTerminationToken && token is not EofToken)
            {
                throw VBSyntaxError(code);
            }
        }

        private void ExpectPunctuation(Punctuation type, VBSyntaxErrorCode code = VBSyntaxErrorCode.SyntaxError)
        {
            var token = Move();
            if (!(token is PunctuationToken p && p.Type == type))
            {
                throw VBSyntaxError(code);
            }
        }

        private void ExpectRParen(VBSyntaxErrorCode code = VBSyntaxErrorCode.ExpectedRParen)
            => ExpectPunctuation(Punctuation.RParen, code);

        private int ExpectInteger(VBSyntaxErrorCode code = VBSyntaxErrorCode.ExpectedInteger)
        {
            var token = Move();
            if (token is DecIntegerLiteralToken i)
            {
                return i.Value;
            }
            
            throw VBSyntaxError(code);
        }

        private string ExpectIdentifier(VBSyntaxErrorCode code = VBSyntaxErrorCode.ExpectedIdentifier)
        {
            var token = Move();
            if (token is IdentifierToken id)
            {
                return id.ToString();
            }
            else if (token is KeywordOrIdentifierToken id2)
            {
                return id2.ToString();
            }

            throw VBSyntaxError(code);
        }

        private Identifier CreateIdentifier(string name)
        {
            if (name?.Length > Identifier.MaxLength)
            {
                throw VBSyntaxError(VBSyntaxErrorCode.IdentifierTooLong);
            }

            return new Identifier(name!);
        }

        private Exception VBSyntaxError(VBSyntaxErrorCode code)
        {
            return new VBSyntaxErrorException(
                code, _lastMarker.Line, _lastMarker.Column);
        }

        private void SkipCommentsAndNewlines()
        {
            void SkipNewlines()
            {
                while (_next is LineTerminationToken)
                {
                    Move();
                }
            }

            while (_next is LineTerminationToken || _next is CommentToken)
            {
                SkipComments();
                SkipNewlines();
            }
        }

        private void SkipComments()
        {
            while (_next is CommentToken tk)
            {
                Marker marker = default;
                Comment? comment = null;
                if (_options.SaveComments)
                {
                    marker = CreateMarker();
                    comment = new Comment(tk.IsRem ? CommentType.Rem : CommentType.SingleQuote, tk.Comment);
                    _comments.Add(comment);
                }

                Move();

                if (comment != null)
                {
                    comment.Range = new Range(marker.Index, _lastMarker.Index);
                    var start = new Position(marker.Line, marker.Column);
                    var end = new Position(_lastMarker.Line, _lastMarker.Column);
                    comment.Location = new Location(start, end);
                }
            }
        }

        private Token Move()
        {
            var token = _next;

            _lastMarker.Index = _lexer.Index;
            _lastMarker.Line = _lexer.CurrentLine;
            _lastMarker.Column = _lexer.LineIndex;

            _lexer.SkipWhitespaces();

            if (_lexer.Index != _startMarker.Index)
            {
                _startMarker.Index = _lexer.Index;
                _startMarker.Line = _lexer.CurrentLine;
                _startMarker.Column = _lexer.LineIndex;
            }

            _next = _lexer.NextToken();

            return token;
        }

        private Marker CreateMarker()
            => new Marker(_startMarker.Index, _startMarker.Line, _startMarker.Column);

        private Marker CreateMarker(Token token, int lastLineStart = 0)
        {
            var column = token.Start - token.LineStart;
            var line = token.LineNumber;
            if (column < 0)
            {
                column += lastLineStart;
                line--;
            }
            return new Marker(token.Start, line, column);
        }

        private T Finalize<T>(Marker marker, T node) where T : Node
        {
            node.Range = new Range(marker.Index, _lastMarker.Index);

            var start = new Position(marker.Line, marker.Column);
            var end = new Position(_lastMarker.Line, _lastMarker.Column);
            node.Location = new Location(start, end);

            return node;
        }

        private delegate T VariableDeclarationCtor<T>(Identifier id, bool isDynamicArray)
            where T : VariableDeclarationNode;

        private delegate T SubDeclarationCtor<T>(MethodAccessModifier modifier, Identifier id, Statement body)
            where T : ProcedureDeclaration;

        private delegate PropertyDeclaration PropertyDeclarationCtor(MethodAccessModifier modifier, Identifier id);
    }
}
