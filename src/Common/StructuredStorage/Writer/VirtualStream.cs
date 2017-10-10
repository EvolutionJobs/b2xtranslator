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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Class which represents a virtual stream in a structured storage.
    /// Author: math
    /// </summary>
    internal class VirtualStream
    {
        AbstractFat _fat;
        Stream _stream;
        ushort _sectorSize;
        OutputHandler _outputHander;

        // Start sector of the virtual stream.
        uint _startSector = SectorId.FREESECT;
        public uint StartSector
        {
            get { return this._startSector; }
        }
        
        // Lengh of the virtual stream.
        public ulong Length
        {
            get { return (ulong)this._stream.Length; }
        }

        // Number of sectors used by the virtual stream.
        uint _sectorCount;
        public uint SectorCount
        {
            get { return this._sectorCount;  }
        }


        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="stream">The input stream.</param>
        /// <param name="fat">The fat which is used by this stream.</param>
        /// <param name="sectorSize">The sector size.</param>
        /// <param name="outputHander"></param>
        internal VirtualStream(Stream stream, AbstractFat fat, ushort sectorSize, OutputHandler outputHander)
        {
            this._stream = stream;
            this._fat = fat;
            this._sectorSize = sectorSize;
            this._outputHander = outputHander;
            this._sectorCount = (uint)Math.Ceiling((double)this._stream.Length / (double)this._sectorSize);
        }


        /// <summary>
        /// Writes the virtual stream chain to the fat and the virtual stream data to the output stream of the current context.
        /// </summary>
        internal void write()
        {
            this._startSector = this._fat.writeChain(this.SectorCount);
            var reader = new BinaryReader(this._stream);
            reader.BaseStream.Seek(0, SeekOrigin.Begin);
            while (true) {
                var bytes = reader.ReadBytes((int)this._sectorSize);
                this._outputHander.writeSectors(bytes, this._sectorSize, (byte)0x0);
                if (bytes.Length != this._sectorSize)
                {
                    break;
                }
            }
        }
    }
}
