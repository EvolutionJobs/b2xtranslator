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

using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class FrameSequence : BiffRecordSequence, IVisitable
    {
        public Frame Frame;
        
        public Begin Begin;
        
        public LineFormat LineFormat;
        
        public AreaFormat AreaFormat;
        
        public GelFrameSequence GelFrameSequence;
        
        public ShapePropsSequence ShapePropsSequence;

        public End End;

        public FrameSequence(IStreamReader reader) : base(reader)
        {
            // FRAME = Frame Begin LineFormat AreaFormat [GELFRAME] [SHAPEPROPS] End

            // Frame 
            this.Frame = (Frame)BiffRecord.ReadRecord(reader);
            
            // Begin 
            this.Begin = (Begin)BiffRecord.ReadRecord(reader); 
            
            // LineFormat 
            this.LineFormat = (LineFormat)BiffRecord.ReadRecord(reader);
            
            // AreaFormat 
            this.AreaFormat = (AreaFormat)BiffRecord.ReadRecord(reader);
            
            // [GELFRAME] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.GelFrame)
            {
                this.GelFrameSequence = new GelFrameSequence(reader);
            }
            
            // [SHAPEPROPS] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.ShapePropsStream)
            {
               this.ShapePropsSequence = new ShapePropsSequence(reader);
            }
            
            // End
            this.End = (End)BiffRecord.ReadRecord(reader);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<FrameSequence>)mapping).Apply(this);
        }

        #endregion
    }
}
