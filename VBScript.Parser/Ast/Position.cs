using System;
using System.Globalization;

namespace VBScript.Parser.Ast
{
    public readonly struct Position : IEquatable<Position>
    {
        public Position(int line, int column)
        {
            Line = line;
            Column = column;
        }

        public int Line { get; }
        public int Column { get; }

        public override bool Equals(object obj) =>
            obj is Position other && Equals(other);

        public bool Equals(Position other) =>
            Line == other.Line && Column == other.Column;

        public override int GetHashCode() =>
            unchecked((Line * 397) ^ Column);

        public override string ToString()
            => Line.ToString(CultureInfo.InvariantCulture)
             + ","
             + Column.ToString(CultureInfo.InvariantCulture);

        public static bool operator ==(Position left, Position right) => left.Equals(right);
        public static bool operator !=(Position left, Position right) => !left.Equals(right);

        public void Deconstruct(out int line, out int column)
        {
            line = Line;
            column = Column;
        }
    }
}
