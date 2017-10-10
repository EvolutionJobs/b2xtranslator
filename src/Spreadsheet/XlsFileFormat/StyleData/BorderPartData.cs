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


namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.StyleData
{
    public class BorderPartData
    {
        public ushort style;
        public int colorId;

        public BorderPartData(ushort style, int colorId)
        {
            this.style = style; 
            this.colorId = colorId; 
        }

        /// <summary>
        /// Equals Method 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(System.Object obj)
        {
            // If parameter is null return false.
            if (obj == null)
            {
                return false;
            }

            // If parameter cannot be cast to FillDataList return false.
            var bpd = obj as BorderPartData;
            if ((System.Object)bpd == null)
            {
                return false;
            }

            // Return true if the fields match:
            return (this.colorId == bpd.colorId) && (this.style == bpd.style);
        }

        /// <summary>
        /// Equals Method
        /// </summary>
        /// <param name="fd"></param>
        /// <returns></returns>
        public bool Equals(BorderPartData bpd)
        {
            // If parameter is null return false:
            if ((object)bpd == null)
            {
                return false;
            }

            // Return true if the fields match:
            return (this.colorId == bpd.colorId) && (this.style == bpd.style);
        }
    }


}
