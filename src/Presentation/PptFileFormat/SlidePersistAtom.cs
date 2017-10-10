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
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecord(1011)]
    public class SlidePersistAtom : Record
    {
        /// <summary>
        /// logical reference to the slide persist object
        /// </summary>
        public uint PersistIdRef;

        /// <summary>
        /// Bit 1: Slide outline view is collapsed
        /// Bit 2: Slide contains shapes other than placeholders
        /// </summary>
        public uint Flags;

        /// <summary>
        /// number of placeholder texts stored with the persist object.
        /// Allows to display outline view without loading the slide persist objects
        /// </summary>
        public int NumberText;

        /// <summary>
        /// Unique slide identifier, used for OLE link monikers for example
        /// </summary>
        public uint SlideId;

        public SlidePersistAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.PersistIdRef = this.Reader.ReadUInt32();
            this.Flags = this.Reader.ReadUInt32();
            this.NumberText = this.Reader.ReadInt32();
            this.SlideId = this.Reader.ReadUInt32();
            this.Reader.ReadUInt32(); // Throw away reserved field
        }

        override public string ToString(uint depth)
        {

            return String.Format("{0}\n{1}PsrRef = {2}\n{1}Flags = {3}, NumberText = {4}, SlideId = {5})",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.PersistIdRef, this.Flags, this.NumberText, this.SlideId);
        }
    }
}
