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

using System;
using System.Collections.Generic;
using System.Text;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.Tools
{

    public class EmuValue
    {
        public int Value;

        /// <summary>
        /// Creates a new EmuValue for the given value.
        /// </summary>
        /// <param name="value"></param>
        public EmuValue(int value)
        {
            this.Value = value;
        }

        /// <summary>
        /// Converts the EMU to pt
        /// </summary>
        /// <returns></returns>
        public double ToPoints()
        {
            return this.Value / 12700;
        }

        /// <summary>
        /// Converts the EMU to twips
        /// </summary>
        /// <returns></returns>
        public double ToTwips()
        {
            return this.Value / 635;
        }

        public double ToCm()
        {
            return this.Value / 36000.0;
        }

        public double ToMm()
        {
            return ToCm() * 10.0;
        }

        /// <summary>
        /// returns the original value as string 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Convert.ToString(this.Value, CultureInfo.GetCultureInfo("en-US")); 
        }
    }
}
