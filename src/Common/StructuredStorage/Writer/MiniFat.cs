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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Represents the minifat of a structured storage.
    /// Author: math
    /// </summary>
    internal class MiniFat : AbstractFat
    {
        // Start sector of the minifat.
        uint _miniFatStart = SectorId.FREESECT;
        internal uint MiniFatStart
        {
            get { return _miniFatStart; }            
        }


        // Number of sectors in the mini fat.
        uint _numMiniFatSectors = 0x0;
        internal uint NumMiniFatSectors
        {
            get { return _numMiniFatSectors; }            
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="context">the current context</param>
        internal MiniFat(StructuredStorageContext context)
            : base(context)
        {            
        }


        /// <summary>
        /// Writes minifat chain to fat and writes the minifat data to the output stream of the current context.
        /// </summary>
        override internal void write()
        {
            _numMiniFatSectors = (uint)Math.Ceiling((double)(_entries.Count * 4) / (double)_context.Header.SectorSize);
            _miniFatStart = _context.Fat.writeChain(_numMiniFatSectors);
         
            _context.TempOutputStream.writeSectors(_context.InternalBitConverter.getBytes(_entries).ToArray(), _context.Header.SectorSize, SectorId.FREESECT);
        }

    }
}
