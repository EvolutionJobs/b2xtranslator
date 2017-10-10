/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */


using System;
using System.Runtime.Serialization;

/// <summary>
/// Exceptions used
/// Author: math
/// </summary>

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Common
{
    [Serializable]
    public class MagicNumberException : Exception
    {
        public MagicNumberException()
            : base("Magic Number not found in file.")
        {
        }

        public MagicNumberException(string additionalMessage)
            : base("Magic Number not found in file. " + additionalMessage)
        {
        }

        protected MagicNumberException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class ValueNotZeroException : Exception
    {
        public ValueNotZeroException(string value)
            : base(value + " must be zero.")
        {
        }

        protected ValueNotZeroException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class ReadBytesAmountMismatchException : Exception
    {
        public ReadBytesAmountMismatchException()
            : base("The number of bytes read mismatches the specified amount.")
        {
        }

        protected ReadBytesAmountMismatchException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class UnsupportedSizeException : Exception
    {
        public UnsupportedSizeException(string value)
            : base("The size of " + value + " is not supported.")
        {
        }

        protected UnsupportedSizeException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class InvalidValueInHeaderException : Exception
    {
        public InvalidValueInHeaderException(string value)
            : base("The value for '" + value + "' in the header is invalid.")
        {
        }

        public InvalidValueInHeaderException(string value, string additionalMessage)
            : base("The value for '" + value + "' in the header is invalid. " + additionalMessage)
        {
        }

        protected InvalidValueInHeaderException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class ChainCycleDetectedException : Exception
    {
        public ChainCycleDetectedException(string chain)
            : base(chain + " contains a cycle.")
        {
        }

        protected ChainCycleDetectedException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class ChainSizeMismatchException : Exception
    {
        public ChainSizeMismatchException(string name)
            : base("The number of sectors used by " + name + " does not match the specified size.")
        {
        }

        public ChainSizeMismatchException(string name, string additionalMessage)
            : base("The number of sectors used by " + name + " does not match the specified size. " + additionalMessage)
        {
        }

        protected ChainSizeMismatchException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class InvalidSectorInChainException : Exception
    {
        public InvalidSectorInChainException()
            : base("Chain could not be build due to an invalid sector id.")
        {
        }

        protected InvalidSectorInChainException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class StreamNotInitalizedException : Exception
    {
        public StreamNotInitalizedException()
            : base("The current stream is not initialized.")
        {
        }

        protected StreamNotInitalizedException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class InvalidValueInDirectoryEntryException : Exception
    {
        public InvalidValueInDirectoryEntryException(string value)
            : base("The value for '" + value + "' is invalid.")
        {
        }

        protected InvalidValueInDirectoryEntryException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class WrongDirectoryEntryTypeException : Exception
    {
        public WrongDirectoryEntryTypeException()
            : base("The directory entry is not of type STGTY_STREAM.")
        {
        }

        protected WrongDirectoryEntryTypeException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class StreamNotFoundException : Exception
    {
        public StreamNotFoundException(string name)
            : base("Stream with name '" + name + "' not found.")
        {
        }

        protected StreamNotFoundException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class DirectoryEntryNotFoundException : Exception
    {
        public DirectoryEntryNotFoundException(string name)
            : base("DirectoryEntry with name '" + name + "' not found.")
        {
        }

        protected DirectoryEntryNotFoundException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class FileHandlerNotCorrectlyInitializedException : Exception
    {
        public FileHandlerNotCorrectlyInitializedException()
            : base("The file handler is not correctly initialized.")
        {
        }

        protected FileHandlerNotCorrectlyInitializedException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class DiFatInconsistentException : Exception
    {
        public DiFatInconsistentException()
            : base("Inconsistancy found while writing DiFat.")
        {
        }

        protected DiFatInconsistentException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

    [Serializable]
    public class InvalidSectorSizeException : Exception
    {
        public InvalidSectorSizeException()
            : base("Inconsistancy found while writing a sector.")
        {
        }

        protected InvalidSectorSizeException(SerializationInfo info, StreamingContext ctxt)
            : base(info, ctxt)
        {
        }
    }

}
