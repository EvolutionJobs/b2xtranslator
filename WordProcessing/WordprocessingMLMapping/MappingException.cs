using System;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class MappingException : Exception
    {
        public MappingException(string message)
            : base(message)
        { }
    }
}
