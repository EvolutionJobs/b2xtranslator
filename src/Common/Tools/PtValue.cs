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

    public class PtValue
    {
        public double Value;

        /// <summary>
        /// Creates a new PtValue for the given value.
        /// </summary>
        /// <param name="value"></param>
        public PtValue(double value)
        {
            this.Value = value;
        }

        /// <summary>
        /// Converts the EMU to pt
        /// </summary>
        /// <returns></returns>
        public double ToPoints()
        {
            return this.Value;
        }

        /// <summary>
        /// Converts the pt value to EMU
        /// </summary>
        /// <returns></returns>
        public int ToEmu()
        {
            return (int)((360000 * 2.54 * this.Value) / 72.0);
        }

        /// <summary>
        /// Converts the pt value to cm
        /// </summary>
        /// <returns></returns>
        public double ToCm()
        {
            return (2.54 * this.Value) / 72.0;
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
