using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;

namespace VBScript.Parser
{
    [Serializable]
    public class VBSyntaxErrorException : Exception
    {
        public VBSyntaxErrorException(VBSyntaxErrorCode code, int line, int position)
            : this(code, line, position, VBSyntaxErrorMessages.ResourceManager.GetString(code))
        {
        }

        public VBSyntaxErrorException(VBSyntaxErrorCode code, int line, int position, string message)
            : this(code, line, position, message, null)
        {
        }

        public VBSyntaxErrorException(VBSyntaxErrorCode code, int line, int position, string message, Exception? innerException)
            : base(message, innerException)
        {
            Code = code;
            Line = line;
            Position = position;
        }

        protected VBSyntaxErrorException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        public VBSyntaxErrorCode Code { get; }
        public int Line { get; }
        public int Position { get; }
    }
}
