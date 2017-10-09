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
    public class FillData : System.Object
    {
        /// <summary>
        /// Type from this filldata object 
        /// </summary>
        private StyleEnum fillPatern;
        public StyleEnum Fillpatern
        {
            get { return this.fillPatern; }
        }

        /// <summary>
        /// Foreground Color 
        /// </summary>
        private int icvFore;
        public int IcvFore
        {
            get { return icvFore; }
        }

        /// <summary>
        /// Background color 
        /// </summary>
        private int icvBack;
        public int IcvBack
        {
            get { return icvBack; }
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="fillpat">Fill Patern</param>
        /// <param name="icvFore">Foreground Color</param>
        /// <param name="icvBack">Background Color</param>
        public FillData(StyleEnum fillpat, int icvFore, int icvBack)
        {
            this.fillPatern = fillpat;
            this.icvFore = icvFore;
            this.icvBack = icvBack;
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
            FillData fd = obj as FillData;
            if ((System.Object)fd == null)
            {
                return false;
            }

            // Return true if the fields match:
            return (this.fillPatern == fd.fillPatern) && (this.icvBack == fd.icvBack) && (this.icvFore == fd.icvFore);
        }

        /// <summary>
        /// Equals Method
        /// </summary>
        /// <param name="fd"></param>
        /// <returns></returns>
        public bool Equals(FillData fd)
        {
            // If parameter is null return false:
            if ((object)fd == null)
            {
                return false;
            }

            // Return true if the fields match:
            return (this.fillPatern == fd.fillPatern) && (this.icvBack == fd.icvBack) && (this.icvFore == fd.icvFore);
        }

        /// <summary>
        /// Simple toString method
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return "Fillvalue: " + this.fillPatern.ToString() + "   FG: " + this.icvFore.ToString() + "  BG: " + this.icvBack.ToString();
        }

    }
}
