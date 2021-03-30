using System;
using System.Collections.Generic;
using System.Text;
using VBScript.Parser.Ast;

namespace VBScript.Parser
{
    public abstract class Visitor<T>
    {
        public abstract T Visit(Node node);

        public virtual T Visit(Program node) => Visit(node as Node);
        public virtual T Visit(Parameter node) => Visit(node as Node);
        public virtual T Visit(ConstDeclaration node) => Visit(node as Node);
        public virtual T Visit(ReDimDeclaration node) => Visit(node as Node);

        public virtual T Visit(VariableDeclarationNode node) => Visit(node as Node);
        public virtual T Visit(VariableDeclaration node) => Visit(node as VariableDeclarationNode);
        public virtual T Visit(FieldDeclaration node) => Visit(node as VariableDeclarationNode);

        public virtual T Visit(Statement stmt) => Visit(stmt as Node);

        public virtual T Visit(ProcedureDeclaration stmt) => Visit(stmt as Statement);
        public virtual T Visit(SubDeclaration stmt) => Visit(stmt as ProcedureDeclaration);
        public virtual T Visit(InitializeSubDeclaration stmt) => Visit(stmt as SubDeclaration);
        public virtual T Visit(TerminateSubDeclaration stmt) => Visit(stmt as SubDeclaration);
        public virtual T Visit(FunctionDeclaration stmt) => Visit(stmt as ProcedureDeclaration);

        public virtual T Visit(PropertyDeclaration stmt) => Visit(stmt as Statement);
        public virtual T Visit(PropertyGetDeclaration stmt) => Visit(stmt as PropertyDeclaration);
        public virtual T Visit(PropertySetDeclaration stmt) => Visit(stmt as PropertyDeclaration);
        public virtual T Visit(PropertyLetDeclaration stmt) => Visit(stmt as PropertyDeclaration);

        public virtual T Visit(Expression expr) => Visit(expr as Node);
        
        public virtual T Visit(LiteralExpression expr) => Visit(expr as Expression);
        public virtual T Visit(BooleanLiteral expr) => Visit(expr as LiteralExpression);
        public virtual T Visit(DateLiteral expr) => Visit(expr as LiteralExpression);
        public virtual T Visit(FloatLiteral expr) => Visit(expr as LiteralExpression);
        public virtual T Visit(NullLiteral expr) => Visit(expr as LiteralExpression);
        public virtual T Visit(StringLiteral expr) => Visit(expr as LiteralExpression);
        public virtual T Visit(IntegerLiteral expr) => Visit(expr as LiteralExpression);
        public virtual T Visit(NothingLiteral expr) => Visit(expr as LiteralExpression);
        public virtual T Visit(EmptyLiteral expr) => Visit(expr as LiteralExpression);

        public virtual T Visit(Identifier expr) => Visit(expr as Expression);
        public virtual T Visit(UnaryExpression expr) => Visit(expr as Expression);
        public virtual T Visit(BinaryExpression expr) => Visit(expr as Expression);
        public virtual T Visit(IndexOrCallExpression expr) => Visit(expr as Expression);
        public virtual T Visit(MemberExpression expr) => Visit(expr as Expression);
        public virtual T Visit(MissingValueExpression expr) => Visit(expr as Expression);
        public virtual T Visit(NewExpression expr) => Visit(expr as Expression);
        public virtual T Visit(WithMemberAccessExpression expr) => Visit(expr as Expression);

        public virtual T Visit(AssignmentStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(CallStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(CallSubStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(CaseStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(ClassDeclaration stmt) => Visit(stmt as Statement);
        public virtual T Visit(ConstsDeclaration stmt) => Visit(stmt as Statement);
        public virtual T Visit(DoStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(ElseIfStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(EraseStatement stmt) => Visit(stmt as Statement);

        public virtual T Visit(ExitStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(ExitDoStatement stmt) => Visit(stmt as ExitStatement);
        public virtual T Visit(ExitForStatement stmt) => Visit(stmt as ExitStatement);
        public virtual T Visit(ExitFunctionStatement stmt) => Visit(stmt as ExitStatement);
        public virtual T Visit(ExitPropertyStatement stmt) => Visit(stmt as ExitStatement);
        public virtual T Visit(ExitSubStatement stmt) => Visit(stmt as ExitStatement);

        public virtual T Visit(FieldsDeclaration stmt) => Visit(stmt as Statement);
        public virtual T Visit(ForEachStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(ForStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(IfStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(OnErrorGoTo0Statement stmt) => Visit(stmt as Statement);
        public virtual T Visit(OnErrorResumeNextStatement stmt) => Visit(stmt as Statement);
        
        public virtual T Visit(ReDimStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(SelectStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(StatementList stmt) => Visit(stmt as Statement);
        
        public virtual T Visit(VariablesDeclaration stmt) => Visit(stmt as Statement);
        public virtual T Visit(WhileStatement stmt) => Visit(stmt as Statement);
        public virtual T Visit(WithStatement stmt) => Visit(stmt as Statement);
    }
}
