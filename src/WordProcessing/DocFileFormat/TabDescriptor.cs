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

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class TabDescriptor
    {
        /// <summary>
        /// Justification code:<br/>
        /// 0 left tab<br/>
        /// 1 centered tab<br/>
        /// 2 right tab<br/>
        /// 3 decimal tab<br/>
        /// 4 bar
        /// </summary>
        public byte jc;

        /// <summary>
        /// Tab leader code:<br/>
        /// 0 no leader<br/>
        /// 1 dotted leader<br/>
        /// 2 hyphenated leader<br/>
        /// 3 single line leader<br/>
        /// 4 heavy line leader<br/>
        /// 5 middle dot
        /// </summary>
        public byte tlc;

        /// <summary>
        /// Parses the bytes to retrieve a TabDescriptor
        /// </summary>
        /// <param name="b">The byte</param>
        public TabDescriptor(byte b)
        {
          this.jc = Convert.ToByte(Convert.ToInt32(b) & 0x07);
          this.tlc = Convert.ToByte(b >> 3);
        }
    }
}
