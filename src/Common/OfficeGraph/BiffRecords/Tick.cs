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

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    /// <summary>
    /// This record specifies the attributes of the axis labels, major tick marks, and minor tick marks associated with an axis.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Tick)]
    public class Tick : OfficeGraphBiffRecord
    {
        public enum MarkLocation
        {
            None,
            Inside,
            Outside,
            Crossing
        }

        public enum MarkLabelLocation
        {
            Complex,
            Low,
            High,
            NextToAxis
        }

        public const GraphRecordNumber ID = GraphRecordNumber.Tick;

        /// <summary>
        /// Specifies the location of major tick marks.
        /// </summary>
        public MarkLocation tktMajor;

        /// <summary>
        /// Specifies the location of minor tick marks.
        /// </summary>
        public MarkLocation tktMinor;

        /// <summary>
        /// Specifies the location of axis labels.
        /// </summary>
        public MarkLabelLocation tlt;

        /// <summary>
        /// Specifies the display mode of the background of the text of the axis labels.
        /// </summary>
        public BackgroundMode wBkgMode;

        /// <summary>
        /// Specifies the color of the text for the axis labels.
        /// </summary>
        public RGBColor rgb;

        /// <summary>
        /// A bit that specifies if the foreground text color of the axis labels is determined automatically.
        /// </summary>
        public bool fAutoCo;

        /// <summary>
        /// A bit that specifies if the background color of the axis label is determined automatically.
        /// </summary>
        public bool fAutoMode;

        /// <summary>
        /// Specifies text rotation of the axis labels. <br/>
        /// If Custom use value from trot.
        /// </summary>
        public TextRotation rot;

        /// <summary>
        /// A bit that specifies whether the text rotation of the axis label is determined automatically.
        /// </summary>
        public bool fAutoRot;

        /// <summary>
        /// specifies the reading order of the axis labels.
        /// </summary>
        public ReadingOrder iReadingOrder;

        /// <summary>
        /// Specifies the color of the text in the color palette.
        /// </summary>
        public UInt16 icv;

        /// <summary>
        /// MUST be a value from the following table:<br/>
        /// 0 to 90 = Text rotated 0 to 90 degrees counterclockwise<br/>
        /// 91 to 180 = Text rotated 1 to 90 degrees clockwise (angle is trot – 90)<br/>
        /// 255 = Text top-to-bottom with letters upright
        /// </summary>
        public UInt16 trot;

        public Tick(IStreamReader reader, GraphRecordNumber id, UInt16 length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.tktMajor = (MarkLocation)reader.ReadByte();
            this.tktMinor = (MarkLocation)reader.ReadByte();
            this.tlt = (MarkLabelLocation)reader.ReadByte();
            this.wBkgMode = (BackgroundMode)reader.ReadByte();
            this.rgb = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);
            reader.ReadBytes(16); // rerserved
            UInt16 flags = reader.ReadUInt16();
            this.fAutoCo = Utils.BitmaskToBool(flags, 0x1);
            this.fAutoMode = Utils.BitmaskToBool(flags, 0x2);
            this.rot = (TextRotation)Utils.BitmaskToInt(flags, 0x1C);
            this.fAutoRot = Utils.BitmaskToBool(flags, 0x10);
            this.iReadingOrder = (ReadingOrder)Utils.BitmaskToInt(flags, 0xC000);
            this.icv = reader.ReadUInt16();
            this.trot = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
