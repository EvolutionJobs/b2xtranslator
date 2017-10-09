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
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Reader
{
    public class VirtualStreamReader : BinaryReader, IStreamReader
    {
        /// <summary>
        /// Ctor 
        /// 
        /// </summary>
        /// <param name="stream"></param>
        public VirtualStreamReader(VirtualStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Second constructor to create a StreamReader with a MemoryStream. 
        /// </summary>
        /// <param name="stream"></param>
        public VirtualStreamReader(MemoryStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Reads bytes from the current position in the virtual stream.
        /// The number of bytes to read is determined by the length of the array.
        /// </summary>
        /// <param name="buffer">Array which will contain the read bytes after successful execution.</param>
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the length of the array if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        public int Read(byte[] buffer)
        {
            return BaseStream.Read(buffer, 0, buffer.Length);
        }

        /// <summary>
        /// Reads bytes from the current position in the virtual stream.
        /// </summary>
        /// <param name="buffer">Array which will contain the read bytes after successful execution.</param>
        /// <param name="count">Number of bytes to read.</param>
        /// <returns>The total number of bytes read into the buffer. 
        /// This might be less than the number of bytes requested if that number 
        /// of bytes are not currently available, or zero if the end of the stream is reached.</returns>
        public int Read(byte[] buffer, int count)
        {
            return BaseStream.Read(buffer, 0, count);
        }

        /// <summary>
        /// Reads count bytes from the current stream into a byte array and advances
        ///     the current position by count bytes.
        /// </summary>
        /// <param name="position">The absolute byte offset where to read.</param>
        /// <param name="count">The number of bytes to read.</param>
        /// <returns>A byte array containing data read from the underlying stream. This might
        ///     be less than the number of bytes requested if the end of the stream is reached.</returns>
        /// <exception cref="System.IO.IOException">An I/O error occurs.</exception>
        /// <exception cref="System.ObjectDisposedException">The stream is closed.</exception>
        /// <exception cref="System.ArgumentOutOfRangeException">count is negative.</exception>
        public byte[] ReadBytes(long position, int count)
        {
            BaseStream.Seek(position, SeekOrigin.Begin);
            return ReadBytes(count);
        }
    }
}
