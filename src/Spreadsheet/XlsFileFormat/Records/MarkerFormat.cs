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
    /// This record specifies the color, size, and shape of the associated data markers that 
    /// appear on line, radar, and scatter chart groups. The associated data markers are specified 
    /// by the preceding DataFormat record. If this record is not present in the sequence of records 
    /// that conforms to the SS rule then all the fields will have default values otherwise all 
    /// the fields MUST contain a value.
    /// </summary>
    [BiffRecordAttribute(RecordType.MarkerFormat)]
    public class MarkerFormat : BiffRecord
    {
        public enum MarkerType
        {
            NoMarker,
            SquareMarkers,
            DiamondShapedMarkers,
            TriangularMarkers,
            SquareMarkersWithX,
            SquareMarkersWithAsterisk,
            ShortBarMarkers,
            LongBarMarkers,
            CircularMarkers,
            SquareMarkersWithPlus
        }

        public const RecordType ID = RecordType.MarkerFormat;

        /// <summary>
        /// Specifies the border color of the data marker. <br/>
        /// The color MUST match the color specified by icvFore. <br/>
        /// The default value of this field is automatically selected from the next available color in the Chart color table
        /// </summary>
        public RGBColor rgbFore;

        /// <summary>
        /// Specifies the interior color of the data marker.<br/>
        /// The color MUST match the color specified by icvBack. <br/>
        /// The default value of this field is the same as the default value 
        /// for rgbFore only when the default imk is 0x0001, 0x0002, 0x0003, or 0x0008 otherwise the default value is 0xFFFFFF.
        /// </summary>
        public RGBColor rgbBack;

        /// <summary>
        /// An unsigned integer that specifies the type of data marker. 
        /// </summary>
        public MarkerType imk;

        /// <summary>
        /// A bit that specifies whether the data marker is automatically generated.
        /// </summary>
        public bool fAuto;

        /// <summary>
        /// A bit that specifies whether to show the data marker interior.
        /// </summary>
        public bool fNotShowInt;

        /// <summary>
        /// A bit that specifies whether to show the data marker border.
        /// </summary>
        public bool fNotShowBrd;

        /// <summary>
        /// An unsigned integer that specifies the border color of the data marker.<br/>
        /// The value SHOULD <58> be an IcvChart value. <br/>
        /// The value MUST be an IcvChart value, between 0x0000 and 0x0007 (inclusively), or between 0x0040 and 0x0041 (inclusively).  <br/>
        /// The color MUST match the color specified by rgbFore.  <br/>
        /// The default value of this field is automatically set to match the color specified by rgbFore.
        /// </summary>
        public ushort icvFore;

        /// <summary>
        /// An unsigned integer that specifies the interior color of the data marker.<br/>
        /// The value SHOULD <59> be an IcvChart value. <br/>
        /// The value MUST be an IcvChart value, between 0x0000 and 0x0007 (inclusively), or between 0x0040 and 0x0041 (inclusively).<br/> 
        /// The color MUST match the color specified by rgbBack. <br/>
        /// The default value of this field is automatically set to match the color specified by rgbBack.
        /// </summary>
        public ushort icvBack;

        /// <summary>
        /// An unsigned integer that specifies the size in twips of the data marker. <br/>
        /// MUST be greater than or equal to 40 and less than or equal to 1440. <br/>
        /// The default value for this field is 100.
        /// </summary>
        public uint miSize;

        public MarkerFormat(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.rgbFore = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);
            this.rgbBack = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);
            this.imk = (MarkerType)reader.ReadUInt16();
            var flags = reader.ReadUInt16();
            this.fAuto = Utils.BitmaskToBool(flags, 0x1);
            //0x2 - 0x8 are reserved
            this.fNotShowInt = Utils.BitmaskToBool(flags, 0x10);
            this.fNotShowBrd = Utils.BitmaskToBool(flags, 0x20);
            this.icvFore = reader.ReadUInt16();
            this.icvBack = reader.ReadUInt16();
            this.miSize = reader.ReadUInt32();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
