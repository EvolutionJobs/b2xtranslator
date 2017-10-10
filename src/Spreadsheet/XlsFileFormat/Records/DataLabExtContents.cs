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
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the contents of an extended data label.
    /// </summary>
    [BiffRecordAttribute(RecordType.DataLabExtContents)]
    public class DataLabExtContents : BiffRecord
    {
        public const RecordType ID = RecordType.DataLabExtContents;

        /// <summary>
        /// An FrtHeader. The FrtHeader.rt field MUST be 0x086B.
        /// </summary>
        public FrtHeader frtHeader;

        /// <summary>
        /// A bit that specifies whether the name of the series is displayed in the extended data label.
        /// </summary>
        public bool fSerName;

        /// <summary>
        /// fCatName (1 bit): A bit that specifies whether the category (3) name, or the horizontal 
        /// value on bubble or scatter chart groups, is displayed in the extended data label. 
        /// 
        /// MUST be a value from the following table: 
        ///     Value   Meaning
        ///     0       Neither of the data values are displayed in the extended data label.
        ///     1       If bubble or scatter chart group, the horizontal value is displayed 
        ///             in the extended data label. Otherwise, the category (3) name is 
        ///             displayed in the extended data label.
        /// </summary>
        public bool fCatName;

        /// <summary>
        /// A bit that specifies whether the data value, or the vertical value on bubble or scatter 
        /// chart groups, is displayed in the extended data label. MUST be a value from the following table: 
        /// 
        ///     Value   Meaning
        ///     0       Neither of the data values are displayed in the data label.
        ///     1       If bubble or scatter chart group, the vertical value is displayed 
        ///             in the extended data label. Otherwise, the data value is displayed 
        ///             in the extended data label.
        /// </summary>
        public bool fValue;

        /// <summary>
        /// A bit that specifies whether the value of the corresponding data point, represented as a 
        /// percentage of the sum of the values of the series the data label is associated with, is 
        /// displayed in the extended data label.
        /// 
        /// MUST equal 0 if the chart group type of the corresponding chart group, series, 
        /// or data point is not a bar of pie, doughnut, pie, or pie of pie chart group.
        /// </summary>
        public bool fPercent;

        /// <summary>
        /// A bit that specifies whether the bubble size is displayed in the data label.
        /// 
        /// MUST equal 0 if the chart group type of the corresponding chart group, 
        /// series, or data point is not a bubble chart group.
        /// </summary>
        public bool fBubSizes;

        /// <summary> 
        /// A case-sensitive XLUnicodeStringMin2 that specifies the string that is
        /// inserted between every data value to form the extended data label. For example, if
        /// fCatName and fValue are set to 1, the labels will look like “Category Name<value of
        /// rgchSep>Data Value”. The length of the string is contained in the 
        /// cch field of the XLUnicodeStringMin2 structure. 
        /// </summary>
        public XLUnicodeStringMin2 rgchSep;

        public DataLabExtContents(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeader = new FrtHeader(reader);

            var flags = reader.ReadUInt16();
            this.fSerName = Utils.BitmaskToBool(flags, 0x0001);
            this.fCatName = Utils.BitmaskToBool(flags, 0x0002);
            this.fValue = Utils.BitmaskToBool(flags, 0x0004);
            this.fPercent = Utils.BitmaskToBool(flags, 0x0008);
            this.fBubSizes = Utils.BitmaskToBool(flags, 0x0010);

            this.rgchSep = new XLUnicodeStringMin2(reader);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
