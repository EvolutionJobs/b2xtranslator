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
using System.Collections;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class DateAndTime : IVisitable
    {
        /// <summary>
        /// minutes (0-59)
        /// </summary>
        public short mint;

        /// <summary>
        /// hours (0-23)
        /// </summary>
        public short hr;

        /// <summary>
        /// day of month (1-31)
        /// </summary>
        public short dom;

        /// <summary>
        /// month (1-12)
        /// </summary>
        public short mon;

        /// <summary>
        /// year (1900-2411)-1900
        /// </summary>
        public short yr;

        /// <summary>
        /// weekday<br/>
        /// 0 Sunday
        /// 1 Monday
        /// 2 Tuesday
        /// 3 Wednesday
        /// 4 Thursday
        /// 5 Friday
        /// 6 Saturday
        /// </summary>
        public short wdy;

        /// <summary>
        /// Creates a new DateAndTime with default values
        /// </summary>
        public DateAndTime()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the byte sto retrieve a DateAndTime
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public DateAndTime(byte[] bytes)
        {
            if (bytes.Length == 4)
            {
                var bits = new BitArray(bytes);

                this.mint = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 0, 6));
                this.hr = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 6, 5));
                this.dom = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 11, 5));
                this.mon = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 16, 4));
                this.yr = (short)(1900 + Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 20, 9)));
                this.wdy = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 29, 3));
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct DTTM, the length of the struct doesn't match");
            }
        }

        public DateTime ToDateTime()
        {
            if (this.yr == 1900 && this.mon == 0 && this.dom == 0 && this.hr == 0 && this.mint == 0)
            {
                return new DateTime(1900, 1, 1, 0, 0, 0);
            }
            else
            {
                return new DateTime(this.yr, this.mon, this.dom, this.hr, this.mint, 0);
            } 
        }

        private void setDefaultValues()
        {
            this.dom = 0;
            this.hr = 0;
            this.mint = 0;
            this.mon = 0;
            this.wdy = 0;
            this.yr = 0;
        }

        #region IVisitable Members

        public virtual void Convert<T>(T mapping)
        {
            ((IMapping<DateAndTime>)mapping).Apply(this);
        }

        #endregion
    }
}
