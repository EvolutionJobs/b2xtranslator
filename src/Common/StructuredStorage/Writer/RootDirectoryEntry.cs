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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{
    /// <summary>
    /// Class which represents the root directory entry of a structured storage.
    /// Author: math
    /// </summary>
    public class RootDirectoryEntry : StorageDirectoryEntry
    {
        // The mini stream.
        OutputHandler _miniStream = new OutputHandler(new MemoryStream());
        internal OutputHandler MiniStream
        {
            get { return _miniStream; }
        }


        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="context">the current context</param>
        internal RootDirectoryEntry(StructuredStorageContext context)
            : base("Root Entry", context)
        {            
            Type = DirectoryEntryType.STGTY_ROOT;
            Sid = 0x0;
        }


        /// <summary>
        /// Writes the mini stream chain to the fat and the mini stream data to the output stream of the current context.
        /// </summary>
        internal override void writeReferencedStream()
        {
            var virtualMiniStream = new VirtualStream(_miniStream.BaseStream, Context.Fat, Context.Header.SectorSize, Context.TempOutputStream);
            virtualMiniStream.write();
            this.StartSector = virtualMiniStream.StartSector;
            this.SizeOfStream = virtualMiniStream.Length;
        }

    }
}
