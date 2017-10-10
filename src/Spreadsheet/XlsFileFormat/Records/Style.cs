/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// STYLE: Style Information (293h)
    /// 
    /// Each style in an Excel workbook, whether built-in or user-defined, requires a style 
    /// record in the BIFF file. When Excel saves the workbook, it writes the STYLE records 
    /// in alphabetical order, which is the order in which the styles appear in the drop-down list box.
    /// </summary>
    [BiffRecordAttribute(RecordType.Style)] 
    public class Style : BiffRecord
    {
        public const RecordType ID = RecordType.Style;

        /// <summary>
        /// Index to the style XF record.
        /// 
        /// Note: ixfe uses only the low-order 12 bits of the field (bits 11–0). 
        /// Bits 12, 13, and 14 are unused, and bit 15 ( fBuiltIn ) is 1 for built-in styles.
        /// </summary>
        public ushort ixfe;
                                
        /// <summary>
        /// Built-in style numbers:
        ///     =00h Normal 
        ///     =01h RowLevel_n
        ///     =02h ColLevel_n
        ///     =03h Comma
        ///     =04h Currency
        ///     =05h Percent
        ///     =06h Comma[0]
        ///     =07h Currency[0] 
        /// 
        /// The automatic outline styles — RowLevel_1 through RowLevel_7, 
        /// and ColLevel_1 through ColLevel_7 — are stored by setting istyBuiltIn to 01h or 02h 
        /// and then setting iLevel to the style level minus 1. 
        /// If the style is not an automatic outline style, ignore this field. 
        /// 
        /// Note: for built-in styles only
        /// </summary>
        public byte istyBuiltIn;

        /// <summary>
        /// Level of the outline style RowLevel_n or ColLevel_n (see text) (for built-in styles only).
        /// </summary>
        public byte iLevel;

        /// <summary>
        /// Length of the style name (for non-built-in styles only).
        /// </summary>
        public int cch;

        /// <summary>
        /// Style name (for non-built-in styles only).
        /// </summary>
        public string rgch;

        /// <summary>
        /// A flag indicating whether this is a built-in style
        /// </summary>
        public bool fBuiltin;
 
        public Style(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.ixfe = reader.ReadUInt16();
            this.fBuiltin = Utils.BitmaskToBool(this.ixfe, 0x8000);

            // only bit 11-0 are used
            // TODO: check if the filtering mask need to be applied or not
            this.ixfe = (ushort)(this.ixfe & 0x0FFF);

            if (this.fBuiltin)
            {
                this.istyBuiltIn = reader.ReadByte();
                this.iLevel = reader.ReadByte();
            }
            else
            {
                this.cch = reader.ReadUInt16();
                int grbit = (int)reader.ReadByte();
                this.rgch = ExcelHelperClass.getStringFromBiffRecord(reader, this.cch, grbit); 
            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
