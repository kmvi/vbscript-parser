namespace VBScript.Parser.Ast
{
    public abstract class Node
    {
        public Range Range { get; set; }
        public Location Location { get; set; }
    }
}
