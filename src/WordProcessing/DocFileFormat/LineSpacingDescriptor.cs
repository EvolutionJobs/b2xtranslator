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
    public class LineSpacingDescriptor
    {
        /// <summary>
        /// 
        /// </summary>
        public short dyaLine;

        /// <summary>
        /// 
        /// </summary>
        public bool fMultLinespace;

        /// <summary>
        /// Creates a new LineSpacingDescriptor with empty values
        /// </summary>
        public LineSpacingDescriptor()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a LineSpacingDescriptor
        /// </summary>
        /// <param name="bytes"></param>
        public LineSpacingDescriptor(byte[] bytes)
        {
            if (bytes.Length == 4)
            {
                this.dyaLine = System.BitConverter.ToInt16(bytes, 0);

                if (System.BitConverter.ToInt16(bytes, 2) == 1)
                {
                    this.fMultLinespace = true;
                }
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct LSPD, the length of the struct doesn't match");
            }
        }

        private void setDefaultValues()
        {
            this.dyaLine = 0;
            this.fMultLinespace = false;
        }
    }
}
