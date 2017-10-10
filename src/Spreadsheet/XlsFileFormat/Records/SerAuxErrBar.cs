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
    /// This record specifies properties of an error bar.
    /// </summary>
    [BiffRecordAttribute(RecordType.SerAuxErrBar)]
    public class SerAuxErrBar : BiffRecord
    {
        public const RecordType ID = RecordType.SerAuxErrBar;

        public enum ErrorBarDirection
        {
            HorizontalPositive = 1,
            HorizontalNegative = 2,
            VerticalPositive = 3,
            VerticalNegative = 4
        }

        public enum ErrorAmoutType
        {
            Percentage = 1,
            FixedValue = 2,
            StandardDeviation = 3,
            StandardError = 5
        }

        /// <summary>
        /// Specifies the direction of the error bars.
        /// </summary>
        public ErrorBarDirection sertm;

        /// <summary>
        /// Specifies the error amount type of the error bars.
        /// </summary>
        public ErrorAmoutType ebsrc;

        /// <summary>
        /// A Boolean that specifies whether the error bars are T-shaped.
        /// </summary>
        public bool fTeeTop;

        /// <summary>
        /// An Xnum that specifies the fixed value, percentage, or number of standard deviations for the error bars. 
        /// If ebsrc is StandardError this MUST be ignored.
        /// </summary>
        public double numValue;

        public SerAuxErrBar(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.sertm = (ErrorBarDirection)reader.ReadByte();
            this.ebsrc = (ErrorAmoutType)reader.ReadByte();
            this.fTeeTop = Utils.ByteToBool(reader.ReadByte());
            reader.ReadByte(); // reserved
            this.numValue = reader.ReadDouble();
            reader.ReadBytes(2); //unused

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
