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
    /// <summary>
    /// This record specifies the layout of a picture attached to a picture-filled chart element.
    /// </summary>
    [BiffRecordAttribute(RecordType.PicF)]
    public class PicF : BiffRecord
    {
        public enum LayoutType
        {
            Stretched = 1,
            Stacked = 2,
            StackedAndScaled = 3
        }

        public const RecordType ID = RecordType.PicF;

        /// <summary>
        /// An unsigned integer that specifies the picture layout. 
        /// If this record is not located in the sequence of records that conforms to the SS rule, 
        /// then this value MUST be Stretched.
        /// </summary>
        public LayoutType ptyp;

        /// <summary>
        /// A bit that specifies whether the picture covers the top and bottom fill areas of the data points. 
        /// The top and bottom fill areas of the data points are parallel to the floor in a 3-D plot area. 
        /// If a Chart3d record does not exist in the chart sheet substream, or if this record is not in a 
        /// sequence of records that conforms to the SS rule or if this record is in an SS rule that contains a 
        /// Chart3DBarShape with the riser field equal to 0x01, this field MUST be true.
        /// </summary>
        public bool fTopBottom;

        /// <summary>
        /// A bit that specifies whether the picture covers the front and back fill areas of the data points on 
        /// a bar or column chart group. If a Chart3d record does not exist in the chart sheet substream, 
        /// or if this record is not in a sequence of records that conforms to the SS rule or if this 
        /// record is in an SS rule that contains a Chart3DBarShape with the riser field equal to 0x01, 
        /// this field MUST be true.
        /// </summary>
        public bool fBackFront;

        /// <summary>
        /// A bit that specifies whether the picture covers the side fill areas of the data points on a 
        /// bar or column chart group. If a Chart3d record does not exist in the chart sheet substream, 
        /// or if this record is not in a sequence of records that conforms to the SS rule or if this record 
        /// is in an SS rule that contains a Chart3DBarShape with the riser field equal to 0x01, 
        /// this field MUST be true.
        /// </summary>
        public bool fSide;

        /// <summary>
        /// An Xnum that specifies the number of units on the value axis in which to fit the entire picture.<br/> 
        /// The picture is scaled to fit within this number of units.<br/>
        /// If the value of ptyp is not 0x0003, this field is undefined and MUST be ignored.
        /// </summary>
        public double numScale;

        public PicF(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.ptyp = (LayoutType)reader.ReadInt16();
            reader.ReadBytes(2); //unused
            var flags = reader.ReadUInt16();
            // first 9 bits are reserved
            this.fTopBottom = Utils.BitmaskToBool(flags, 0x200);
            this.fBackFront = Utils.BitmaskToBool(flags, 0x400);
            this.fSide = Utils.BitmaskToBool(flags, 0x800);
            this.numScale = reader.ReadDouble();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
