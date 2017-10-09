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

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// An 8-byte union that specifies a floating-point value, or a non-numeric 
    /// value defined by the containing record. The type and meaning of the union 
    /// contents are determined by the most significant 2 bytes, and is defined in 
    /// the following table: 
    /// 
    ///   Value of most significant 2 bytes       Type and meaning of union contents
    ///     0xFFFF                                  A NilChartNum that specifies a non-numeric value, 
    ///                                             as defined by the containing record.
    ///                                             
    ///     Any other value.                        An Xnum that specifies a floating-point value.
    /// </summary>
    public class ChartNumNillable
    {
        /// <summary>
        /// The nullable Xnum value
        /// </summary>
        public double? value;

        public ChartNumNillable(IStreamReader reader)
        {
            //read the nullable double value 
            byte[] b = reader.ReadBytes(8);
            if (b[6] == 0xFF && b[7] == 0xFF)
            {
                this.value = null;
            }
            else
            {
                this.value = System.BitConverter.ToDouble(b, 0);
            }
        }
    }
}
