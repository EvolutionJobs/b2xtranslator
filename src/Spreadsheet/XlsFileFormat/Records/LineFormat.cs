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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the appearance of a line.
    /// </summary>
    [BiffRecordAttribute(RecordType.LineFormat)]
    public class LineFormat : BiffRecord, IVisitable
    {
        public enum LineStyle
        {
            Solid,
            Dash,
            Dot,
            DashDot,
            DashDotDot,
            None,
            DarkGrayPattern,
            MediumGrayPattern,
            LightGrayPattern
        }

        public enum LineWeight
        {
            Hairline = -1,
            Narrow = 0,
            Medium = 1,
            Wide = 2
        }

        public const RecordType ID = RecordType.LineFormat;

        /// <summary>
        /// Specifies the color of the line.
        /// </summary>
        public RGBColor rgb;

        /// <summary>
        /// Specifies the style of the line.
        /// </summary>
        public LineStyle lns;

        /// <summary>
        /// Specifies the thickness of the line.
        /// </summary>
        public LineWeight we;

        /// <summary>
        /// A bit that specifies whether the line has default formatting.<br/>
        /// If the value is false, the line has formatting as specified by lns, we, and icv.<br/>
        /// If the value is true, lns, we, icv, and rgb MUST be ignored and default values are used as specified in the following table:<br/>
        /// lns = Solid<br/>
        /// we = Narrow<br/>
        /// icv = 0x004D<br/>
        /// rgb = Match the default color used for icv<br/>
        /// </summary>
        public bool fAuto;

        /// <summary>
        /// A bit that specifies whether the axis line is displayed.
        /// </summary>
        public bool fAxisOn;

        /// <summary>
        /// A bit that specifies whether icv equals 0x004D.<br/>
        /// If the value is true, icv MUST equal 0x004D. <br/>
        /// If the value is false, icv MUST NOT equal 0x004D.
        /// </summary>
        public bool fAutoCo;

        /// <summary>
        /// An unsigned integer that specifies the color of the line. 
        /// The value SHOULD be an IcvChart value. <br/>
        /// The value MUST be an IcvChart value, 0x0040, or 0x0041. <br/>
        /// The color MUST match the color specified by rgb.<br/>
        /// </summary>
        public ushort icv;

        public LineFormat(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.rgb = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);
            this.lns = (LineStyle)reader.ReadInt16();
            this.we = (LineWeight)reader.ReadInt16();
            var flags = reader.ReadUInt16();
            this.fAuto = Utils.BitmaskToBool(flags, 0x1);
            // 0x2 is reserved
            this.fAxisOn = Utils.BitmaskToBool(flags, 0x4);
            this.fAutoCo = Utils.BitmaskToBool(flags, 0x8);
            this.icv = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<LineFormat>)mapping).Apply(this);
        }

        #endregion
    }
}
