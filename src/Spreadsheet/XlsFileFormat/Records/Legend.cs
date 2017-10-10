/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using System;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.Legend)]
    public class Legend : BiffRecord
    {
        public const RecordType ID = RecordType.Legend;

        /// <summary>
        /// An unsigned integer that specifies the x-position, in SPRC, 
        /// of the upper-left corner of the bounding rectangle of the legend. <br/>
        /// MUST be ignored and the x1 field from the following Pos record MUST be used instead.
        /// </summary>
        public uint x;

        /// <summary>
        /// An unsigned integer that specifies the y-position, in SPRC, 
        /// of the upper-left corner of the bounding rectangle of the legend. <br/>
        /// MUST be ignored and the y1 field from the following Pos record MUST be used instead.
        /// </summary>
        public uint y;

        /// <summary>
        /// An unsigned integer that specifies the width, in SPRC, of the bounding rectangle of the legend.<br/>
        /// MUST be ignored and the x2 field from the following Pos record MUST be used instead.
        /// </summary>
        public uint dx;

        /// <summary>
        /// An unsigned integer that specifies the height, in SPRC, of the bounding rectangle of the legend.<br/>
        /// MUST be ignored and the y2 field from the following Pos record MUST be used instead.
        /// </summary>
        public uint dy;

        /// <summary>
        /// An unsigned integer that specifies the space between legend entries.<br/>
        /// MUST be 0x01 which represents 40 twips between legend entries.
        /// </summary>
        public byte wSpace;

        /// <summary>
        /// A bit that specifies whether the legend is automatically positioned. 
        /// If this field is true, then fAutoPosX MUST be true and fAutoPosY MUST be true.
        /// </summary>
        public bool fAutoPosition;

        /// <summary>
        /// A bit that specifies whether the x-positioning of the legend is automatic.
        /// </summary>
        public bool fAutoPosX;

        /// <summary>
        /// A bit that specifies whether the y-positioning of the legend is automatic.
        /// </summary>
        public bool fAutoPosY;

        /// <summary>
        /// A bit that specifies the layout of the legend entries. <br/>
        /// True: The legend contains a single column of legend entries.<br/>
        /// False: The legend contains multiple columns of legend entries or the size of the legend has been manually changed from the default size.
        /// </summary>
        public bool fVert;

        /// <summary>
        /// A bit that specifies whether the legend is shown in a data table.
        /// </summary>
        public bool fWasDataTable;

        public Legend(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.x = reader.ReadUInt32();
            this.y = reader.ReadUInt32();
            this.dx = reader.ReadUInt32();
            this.dy = reader.ReadUInt32();
            reader.ReadByte(); // undefined
            this.wSpace = reader.ReadByte();

            var flags = reader.ReadUInt16();
            this.fAutoPosition = Utils.BitmaskToBool(flags, 0x1);
            //0x2 is reserved
            this.fAutoPosX = Utils.BitmaskToBool(flags, 0x4);
            this.fAutoPosY = Utils.BitmaskToBool(flags, 0x8);
            this.fVert = Utils.BitmaskToBool(flags, 0x10);
            this.fWasDataTable = Utils.BitmaskToBool(flags, 0x20);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
