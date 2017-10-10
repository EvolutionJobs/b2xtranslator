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
    /// This record specifies the type, size and position of the frame around a chart 
    /// element as defined by the Chart Sheet Substream ABNF. 
    /// A chart element‟s frame is specified by the Frame record following it.
    /// </summary>
    [BiffRecordAttribute(RecordType.Frame)]
    public class Frame : BiffRecord
    {
        public const RecordType ID = RecordType.Frame;

        public enum FrameStyle : ushort
        {
            /// <summary>
            /// A frame surrounding the chart element.
            /// </summary>
            NoShadow = 0x0000,

            /// <summary>
            /// A frame with shadow surrounding the chart element.
            /// </summary>
            Shadow = 0x0004
        }

        /// <summary>
        /// An unsigned integer that specifies the type of frame to be drawn.
        /// </summary>
        public FrameStyle frt;

        /// <summary>
        /// A bit that specifies if the size of the frame is automatically calculated. 
        /// If the value is 1, the size of the frame is automatically calculated. 
        /// In this case, the width and height specified by the chart element are ignored 
        /// and the size of the frame is calculated automatically. If the value is 0, the 
        /// width and height specified by the chart element are used as the size of the frame.
        /// </summary>
        public bool fAutoSize;

        /// <summary>
        /// A bit that specifies if the position of the frame is automatically calculated. 
        /// If the value is 1, the position of the frame is automatically calculated. 
        /// In this case, the (x, y) specified by the chart element are ignored, and the 
        /// position of the frame is automatically calculated. If the value is 0, the (x, y) 
        /// location specified by the chart element are used as the position of the frame.
        /// </summary>
        public bool fAutoPosition;

        public Frame(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frt = (FrameStyle)reader.ReadUInt16();

            var flags = reader.ReadUInt16();

            this.fAutoSize = Utils.BitmaskToBool(flags, 0x0001);
            this.fAutoPosition = Utils.BitmaskToBool(flags, 0x0002);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
