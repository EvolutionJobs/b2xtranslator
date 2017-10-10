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
using System.IO;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Common
{

    /// <summary>
    /// Abstract class for input and putput handlers.
    /// Author: math
    /// </summary>
    abstract internal class AbstractIOHandler
    {
        protected Stream _stream;
        protected AbstractHeader _header;
        protected InternalBitConverter _bitConverter;

        abstract internal ulong IOStreamSize { get; }

        /// <summary>
        /// Initializes the internal bit converter
        /// </summary>
        /// <param name="isLittleEndian">flag whether big endian or little endian is used</param>
        internal void InitBitConverter(bool isLittleEndian)
        {
            this._bitConverter = new InternalBitConverter(isLittleEndian);
        }

        /// <summary>
        /// Initializes the reference to the header
        /// </summary>
        /// <param name="header"></param>
        internal void SetHeaderReference(AbstractHeader header)
        {
            this._header = header;
        }

        /// <summary>
        /// Closes the file associated with this handler
        /// </summary>
        public void CloseStream()
        {
            if (this._stream != null)
            {
                this._stream.Close();
            }
        }
    }
}
