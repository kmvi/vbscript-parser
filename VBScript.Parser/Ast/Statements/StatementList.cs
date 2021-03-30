using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace VBScript.Parser.Ast
{
    public class StatementList : Statement, IList<Statement>
    {
        private readonly List<Statement> _statements = new();

        public StatementList()
        {

        }

        public Statement this[int index]
        {
            get => _statements[index];
            set => _statements[index] = value;
        }

        public int Count => _statements.Count;

        public bool IsReadOnly => false;

        public void Add(Statement item) => _statements.Add(item);

        public void Clear() => _statements.Clear();

        public bool Contains(Statement item) => _statements.Contains(item);

        public void CopyTo(Statement[] array, int arrayIndex) => _statements.CopyTo(array, arrayIndex);

        public IEnumerator<Statement> GetEnumerator() => _statements.GetEnumerator();

        public int IndexOf(Statement item) => _statements.IndexOf(item);

        public void Insert(int index, Statement item) => _statements.Insert(index, item);

        public bool Remove(Statement item) => _statements.Remove(item);

        public void RemoveAt(int index) => _statements.RemoveAt(index);

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
