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

using System;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies image data for a sheet background.
    /// </summary>
    [BiffRecordAttribute(RecordType.BkHim)]
    public class BkHim : BiffRecord
    {
        public enum ImageFormat
        {
            Bitmap = 0x0009,
            Native = 0x000E
        }

        /// <summary>
        /// Specifies the image format.
        /// </summary>
        public ImageFormat cf;

        /// <summary>
        /// A signed integer that specifies the size of imageBlob in bytes. <br/>
        /// MUST be greater than or equal to 1.
        /// </summary>
        public Int32 lcb;

        /// <summary>
        /// An array of bytes that specifies the image data for the given format.
        /// </summary>
        public byte[] imageBlob;

        public BkHim(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            this.cf = (ImageFormat)reader.ReadInt16();
            reader.ReadBytes(2); // skip 2 bytes
            this.lcb = reader.ReadInt32();
            this.imageBlob = reader.ReadBytes(lcb);
        }
    }
}
