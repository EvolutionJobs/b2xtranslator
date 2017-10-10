/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
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

using System.Collections;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class DrawingObjectGrid
    {
        /// <summary>
        /// x-coordinate of the upper left-hand corner of the grid
        /// </summary>
        public short xaGrid;

        /// <summary>
        /// y-coordinate of the upper left-hand corner of the grid
        /// </summary>
        public short yaGrid;

        /// <summary>
        /// Width of each grid square
        /// </summary>
        public short dxaGrid;

        /// <summary>
        /// Height of each grid square
        /// </summary>
        public short dyaGrid;

        /// <summary>
        /// The number of grid squares (in the y direction) between each 
        /// gridline drawn on the screen. 0 means don‘t display any 
        /// gridlines in the y direction.
        /// </summary>
        public short dyGridDisplay;

        /// <summary>
        /// Suppress display of gridlines
        /// </summary>
        public bool fTurnItOff;

        /// <summary>
        /// The number of grid squares (in the x direction) between each 
        /// gridline drawn on the screen. 0 means don‘t display any 
        /// gridlines in the y direction.
        /// </summary>
        public short dxGridDisplay;

        /// <summary>
        /// If true, the grid will start at the left and top margins and 
        /// ignore xaGrid and yaGrid
        /// </summary>
        public bool fFollowMargins;

        /// <summary>
        /// Parses the bytes to retrieve a DrawingObjectGrid
        /// </summary>
        /// <param name="bytes"></param>
        public DrawingObjectGrid(byte[] bytes)
        {
            if (bytes.Length == 10)
            {
                this.xaGrid = System.BitConverter.ToInt16(bytes, 0);
                this.yaGrid = System.BitConverter.ToInt16(bytes, 2);
                this.dxaGrid = System.BitConverter.ToInt16(bytes, 4);
                this.dyaGrid = System.BitConverter.ToInt16(bytes, 6);

                //split byte 8 and 9 into bits
                var bits = new BitArray(new byte[] { bytes[8], bytes[9] });
                this.dyGridDisplay = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 0, 7));
                this.fTurnItOff = bits[7];
                this.dxGridDisplay = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 8, 7));
                this.fFollowMargins = bits[15];
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct DOGRID, the length of the struct doesn't match");
            }
        }

    }
}
