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
    /// This record specifies a font and font formatting information.
    /// </summary>
    [BiffRecordAttribute(RecordType.Font)]
    public class Font : BiffRecord
    {
        public const RecordType ID = RecordType.Font;

        public enum FontWeight : ushort
        {
            Default = 0,
            Normal = 400,
            Bold = 700
        }

        public enum ScriptStyle : ushort
        {
            NormalScript = 0x0000,
            SuperScript = 0x0001,
            SubScript = 0x0002
        }

        public enum UnderlineStyle : byte
        {
            None = 0x00,
            Single = 0x01,
            Double = 0x02
        }

        public enum FontFamily : byte
        {
            NotApplicable = 0x00,
            Roman = 0x01,
            Swiss = 0x02,
            Modern = 0x03,
            Script = 0x04,
            Decorative = 0x05
        }

        /// <summary>
        /// An unsigned integer that specifies the height of the font in twips. 
        /// 
        /// SHOULD <49> be greater than or equal to 20 and less than or equal to 8191. 
        /// MUST be greater than or equal to 20 and less than 8181, or 0.
        /// </summary>
        public ushort dyHeight;

        /// <summary>
        /// A bit that specifies whether the font is italic.
        /// </summary>
        public bool fItalic;

        /// <summary>
        /// A bit that specifies whether the font has strikethrough formatting applied.
        /// </summary>
        public bool fStrikeOut;

        /// <summary>
        /// A bit that specifies whether the font has an outline effect applied.
        /// </summary>
        public bool fOutline;

        /// <summary>
        /// A bit that specifies whether the font has a shadow effect applied.
        /// </summary>
        public bool fShadow;

        /// <summary>
        /// A bit that specifies whether the font is condensed or not. If the value is 1, the font is condensed.
        /// </summary>
        public bool fCondense;

        /// <summary>
        /// A bit that specifies whether the font is extended or not. If the value is 1, the font is extended.
        /// </summary>
        public bool fExtend;

        /// <summary>
        /// An unsigned integer that specifies the color of the font. 
        /// 
        /// The value SHOULD <50> be an IcvFont value. 
        /// This value MUST be an IcvFont value, or 0.
        /// </summary>
        public ushort icv;
        // TODO: implement IcvFont structure and color mapping

        /// <summary>
        /// An unsigned integer that specifies the font weight. 
        /// 
        /// The value SHOULD <51> be a value from the following table. 
        /// This value MUST be 0, or a value from the following table.
        /// </summary>
        public FontWeight bls;

        /// <summary>
        /// An unsigned integer that specifies whether superscript, subscript, or normal script is used.
        /// </summary>
        public ScriptStyle sss;

        /// <summary>
        /// An unsigned integer that specifies the underline style.
        /// </summary>
        public UnderlineStyle uls;

        /// <summary>
        /// An unsigned integer that specifies the font family, as defined by 
        /// Windows API LOGFONT structure in [MSDN-FONTS]. 
        /// 
        /// MUST be greater than or equal to 0 and less than or equal to 5.
        /// </summary>
        public FontFamily bFamily;

        /// <summary>
        /// An unsigned integer that specifies the character set, as defined 
        /// by Windows API LOGFONT structure in [MSDN-FONTS].
        /// </summary>
        public byte bCharSet;

        /// <summary>
        /// A ShortXLUnicodeString that specifies the name of this font. String 
        /// length MUST be greater than or equal to 1 and less than or equal to 31. 
        /// The fontName.fHighByte field MUST equal 1. MUST NOT contain any null characters.
        /// </summary>
        public ShortXLUnicodeString fontName;

        public Font(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.dyHeight = reader.ReadUInt16();

            var flags = reader.ReadUInt16();

            // 0x0001 is unused
            this.fItalic = Utils.BitmaskToBool(flags, 0x0002);
            // 0x0004 is unused
            this.fStrikeOut = Utils.BitmaskToBool(flags, 0x0008);
            this.fOutline = Utils.BitmaskToBool(flags, 0x0010);
            this.fShadow = Utils.BitmaskToBool(flags, 0x0020);
            this.fCondense = Utils.BitmaskToBool(flags, 0x0040);
            this.fExtend = Utils.BitmaskToBool(flags, 0x0080);

            this.icv = reader.ReadUInt16();
            this.bls = (FontWeight)reader.ReadUInt16();
            this.sss = (ScriptStyle)reader.ReadUInt16();
            this.uls = (UnderlineStyle)reader.ReadByte();
            this.bFamily = (FontFamily)reader.ReadByte();
            this.bCharSet = reader.ReadByte();

            // skip unused byte
            reader.ReadByte();

            this.fontName = new ShortXLUnicodeString(reader);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
