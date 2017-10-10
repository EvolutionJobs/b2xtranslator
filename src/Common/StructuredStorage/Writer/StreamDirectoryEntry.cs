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


using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Represents a stream directory entry in a structured storage.
    /// Author: math
    /// </summary>
    internal class StreamDirectoryEntry : BaseDirectoryEntry
    {
        Stream _stream;


        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="name">Name of the stream directory entry.</param>
        /// <param name="stream">The stream referenced by the stream directory entry.</param>
        /// <param name="context">The current context.</param>
        internal StreamDirectoryEntry(string name, Stream stream, StructuredStorageContext context)
            : base(name, context)
        {
            this._stream = stream;
            this.Type = DirectoryEntryType.STGTY_STREAM;
        }


        /// <summary>
        /// Writes the referenced stream chain to the fat and the referenced stream data to the output stream of the current context.
        /// </summary>
        internal override void writeReferencedStream()
        {
            VirtualStream vStream = null;
            if (this._stream.Length < this.Context.Header.MiniSectorCutoff)
            {
                vStream = new VirtualStream(this._stream, this.Context.MiniFat, this.Context.Header.MiniSectorSize, this.Context.RootDirectoryEntry.MiniStream);
            }
            else
            {
                vStream = new VirtualStream(this._stream, this.Context.Fat, this.Context.Header.SectorSize, this.Context.TempOutputStream);
            }
            vStream.write();
            this.StartSector = vStream.StartSector;
            this.SizeOfStream = vStream.Length;
        }
    }
}
