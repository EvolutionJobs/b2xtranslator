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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class is used  to store the data from a mergecell biffrecord 
    /// </summary>
    public class MergeCellData
    {
        /// <summary>
        /// First row from the merge cell 
        /// </summary>
        public ushort rwFirst;
        /// <summary>
        /// Last row from the merge cell 
        /// </summary>
        public ushort rwLast;
        /// <summary>
        /// First column of the merge cell 
        /// </summary>
        public ushort colFirst;
        /// <summary>
        /// Last colum of the merge cell 
        /// </summary>
        public ushort colLast;

        /// <summary>
        /// Ctor 
        /// </summary>
        public MergeCellData(): this(0,0,0,0) { }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="rwFirst">First row</param>
        /// <param name="rwLast">Last row</param>
        /// <param name="colFirst">First column</param>
        /// <param name="colLast">Last column</param>
        public MergeCellData(ushort rwFirst, ushort rwLast, ushort colFirst, ushort colLast)
        {
            this.rwFirst = rwFirst;
            this.rwLast = rwLast;
            this.colFirst = colFirst;
            this.colLast = colLast; 
        }

        /// <summary>
        /// Converts the classattributes to a string which can be used in the new open xml standard 
        /// This would be: 
        ///     mergeCell ref="B3:C3" 
        ///     ref is  the from the first cell to the last cell 
        /// </summary>
        /// <returns></returns>
        public string getOXMLFormatedData()
        {
            var returnvalue = "";
            returnvalue += ExcelHelperClass.intToABCString(this.colFirst, (this.rwFirst+1).ToString());
            returnvalue += ":";
            returnvalue += ExcelHelperClass.intToABCString(this.colLast, (this.rwLast+1).ToString());
            return returnvalue; 
        }
    }
}
