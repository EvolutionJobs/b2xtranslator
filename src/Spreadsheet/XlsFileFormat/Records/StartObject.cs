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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.StartObject)]
    public class StartObject : BiffRecord
    {
        public const RecordType ID = RecordType.StartObject;

        public FrtHeaderOld frtHeaderOld;

        /// <summary>
        /// An unsigned integer that specifies the type of object that is encompassed by the block.<br/>
        /// MUST be a value from the following table:<br/>
        /// 0x0010<br/>
        /// 0x0011<br/>
        /// 0x0012
        /// </summary>
        public UInt16 iObjectKind;

        /// <summary>
        /// An unsigned integer that specifies the object context.<br/>
        /// MUST be 0x0000.
        /// </summary>
        public UInt16 iObjectContext;

        /// <summary>
        /// An unsigned integer that specifies additional information about the context of the object, 
        /// along with iObjectContext, iObjectInstance2 and iObjectKind. 
        /// This field MUST equal 0x0000 if iObjectKind equals 0x0010 or 0x0012.
        /// MUST be a value from the following table if iObjectKind equals 0x0011:<br/>
        /// 0x0008 = Specifies the application version. &lt;60&gt;<br/>
        /// 0x0009 = Specifies the application version. &lt;61&gt;<br/>
        /// 0x000A = Specifies the application version. &lt;62&gt;<br/>
        /// 0x000B = Specifies the application version. &lt;63&gt;<br/>
        /// 0x000C = Specifies the application version. &lt;64&gt;<br/>
        /// </summary>
        public UInt16 iObjectInstance1;

        /// <summary>
        /// An unsigned integer that specifies more information about the object context, 
        /// along with iObjectContext, iObjectInstance1 and iObjectKind. <br/>
        /// This field MUST equal 0x0000.
        /// </summary>
        public UInt16 iObjectInstance2;

        public StartObject(IStreamReader reader, RecordType id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeaderOld = new FrtHeaderOld(reader);
            this.iObjectKind = reader.ReadUInt16();
            this.iObjectContext = reader.ReadUInt16();
            this.iObjectInstance1 = reader.ReadUInt16();
            this.iObjectInstance2 = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
