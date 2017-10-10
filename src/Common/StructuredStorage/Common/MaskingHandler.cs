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

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Common
{

    /// <summary>
    /// Provides methods for masking/unmasking strings in a path
    /// Author: math
    /// </summary>
    static internal class MaskingHandler
    {
        static readonly UInt32[] CharsToMask = { '%', '\\' };
        

        /// <summary>
        /// Masks the given string
        /// </summary>
        internal static string Mask(string text)
        {
            string result = text;
            foreach (var character in CharsToMask)
	        {
                result = result.Replace(new string((char)character,1), String.Format(CultureInfo.InvariantCulture, "%{0:X4}", character));
	        }
            return result;
        }


        /// <summary>
        /// Unmasks the given string
        /// </summary>
        internal static string UnMask(string text)
        {
            string result = text;
            foreach (var character in CharsToMask)
            {
                result = result.Replace(String.Format("%{0:X4}", character), new string((char)character, 1));
            }
            return result;
        }

    }
}
