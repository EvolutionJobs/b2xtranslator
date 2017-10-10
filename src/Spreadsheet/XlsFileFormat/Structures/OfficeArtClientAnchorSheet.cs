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

using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the anchor position of a drawing object embedded in a sheet.
    /// </summary>
    public class OfficeArtClientAnchorSheet
    {
        /// <summary>
        /// A bit that specifies whether the shape will be kept intact when the cells are moved.
        /// </summary>
        public bool fMove;

        /// <summary>
        /// A bit that specifies whether the shape will be kept intact when the cells are resized. 
        /// 
        /// If fMove is 1, the value MUST be 1.
        /// </summary>
        public bool fSize;

        /// <summary>
        /// A Col256U that specifies the column of the cell under the top left corner of 
        /// the bounding rectangle of the shape.
        /// 
        /// Col256U: An unsigned integer that specifies the zero-based index of a column 
        /// in the sheet that contains this structure. MUST be less than or equal to 0x0100. 
        /// The value 0x0100 specifies that the formatting in the containing record also 
        /// specifies the default column formatting. If additional columns become visible 
        /// at the extreme right of the column range due to column deletion, those columns 
        /// have this default formatting applied.
        /// </summary>
        public ushort colL;

        /// <summary>
        /// A signed integer that specifies the x coordinate of the top left corner of the 
        /// bounding rectangle relative to the corner of the underlying cell. 
        /// The value is expressed as 1024th‘s of that cell‘s width.
        /// </summary>
        public ushort dxL;

        /// <summary>
        /// An unsigned integer that specifies the zero-based index of a row in 
        /// the sheet that contains this structure.
        /// </summary>
        public ushort rwT;

        /// <summary>
        /// A signed integer that specifies the y coordinate of the top left 
        /// corner of the bounding rectangle relative to the corner of the underlying cell. 
        /// The value is expressed as 1024th‘s of that cell‘s height.
        /// </summary>
        public ushort dyT;

        /// <summary>
        /// A Col256U that specifies the column of the cell under the bottom right 
        /// corner of the bounding rectangle of the shape.
        /// </summary>
        public ushort colR;

        /// <summary>
        /// A signed integer that specifies the x coordinate of the bottom right corner
        /// of the bounding rectangle relative to the corner of the underlying cell. 
        /// The value is expressed as 1024th‘s of that cell‘s width.
        /// </summary>
        public ushort dxR;

        /// <summary>
        /// A RwU that specifies the row of the cell under the bottom right 
        /// corner of the bounding rectangle of the shape.
        /// </summary>
        public ushort rwB;

        /// <summary>
        /// A signed integer that specifies the y coordinate of the bottom right corner 
        /// of the bounding rectangle relative to the corner of the underlying cell. 
        /// The value is expressed as 1024th‘s of that cell‘s height.
        /// </summary>
        public ushort dyB;

        public OfficeArtClientAnchorSheet(byte[] rawData)
        {
            this.fMove = Utils.BitmaskToBool(rawData[0], 0x01);
            this.fSize = Utils.BitmaskToBool(rawData[0], 0x02);
            
            this.colL = System.BitConverter.ToUInt16(rawData, 2);
            this.dxL = System.BitConverter.ToUInt16(rawData, 4);
            this.rwT = System.BitConverter.ToUInt16(rawData, 6);
            this.dyT = System.BitConverter.ToUInt16(rawData, 8);
            this.colR = System.BitConverter.ToUInt16(rawData, 10);
            this.dxR = System.BitConverter.ToUInt16(rawData, 12);
            this.rwB = System.BitConverter.ToUInt16(rawData, 14);
            this.dyB = System.BitConverter.ToUInt16(rawData, 16);
        }
    }
}
