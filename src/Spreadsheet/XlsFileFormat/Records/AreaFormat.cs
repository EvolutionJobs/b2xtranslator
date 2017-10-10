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

using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the patterns and colors used in a filled region of a chart. 
    /// 
    /// If this record is not present in the SS rule of the chart sheet substream ABNF 
    /// then all the fields will have default values otherwise all the fields MUST contain a value.
    /// </summary>
    [BiffRecordAttribute(RecordType.AreaFormat)]
    public class AreaFormat : BiffRecord, IVisitable
    {
        public const RecordType ID = RecordType.AreaFormat;

        /// <summary>
        /// A LongRGB that specifies the foreground color of the fill pattern. 
        /// 
        /// The default value of this field is automatically selected from the next available color in the Chart color table.
        /// </summary>
        public RGBColor rgbFore;

        /// <summary>
        /// A LongRGB that specifies the background color of the fill pattern. 
        /// 
        /// The default value of this field is 0xFFFFFF.
        /// </summary>
        public RGBColor rgbBack;

        /// <summary>
        /// An unsigned integer that specifies the type of fill pattern. 
        /// 
        /// If fls is neither 0x0000 nor 0x0001, this record MUST be immediately 
        /// followed by a GelFrame record that specifies the fill pattern. 
        /// 
        /// The fillType as specified in [MS-ODRAW] section 2.3.7.1 of the OPT1 field of the corresponding GelFrame record. 
        /// 
        /// MUST be msofillPattern as specified in [MS-ODRAW] section 2.4.11. The default value of this field is 0x0001.
        /// </summary>
        public ushort fls;

        /// <summary>
        /// A bit that specifies whether the fill colors are automatically set. 
        /// 
        /// If fls equals 0x0001 formatting is automatic. The default value of this field is 1.
        /// </summary>
        public bool fAuto = true;

        /// <summary>
        /// A bit that specifies whether the foreground and background are swapped 
        /// when the data value of the filled area is negative. 
        /// 
        /// This field MUST be ignored if the formatting is not being applied to a data point on a bar or column chart group. 
        /// The default value of this field is 0.
        /// </summary>
        public bool fInvertNeg = false;

        public ushort icvFore;
        public ushort icvBack;

        public AreaFormat(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            rgbFore = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);
            rgbBack = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);

            // TODO: Read optional GelFrame
            fls = reader.ReadUInt16();

            var flags = reader.ReadUInt16();
            fAuto = Utils.BitmaskToBool(flags, 0x1);
            fInvertNeg = Utils.BitmaskToBool(flags, 0x2);

            // TODO: handle default cases and ignoring of fields
            icvFore = reader.ReadUInt16();
            icvBack = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<AreaFormat>)mapping).Apply(this);
        }

        #endregion
    }
}
