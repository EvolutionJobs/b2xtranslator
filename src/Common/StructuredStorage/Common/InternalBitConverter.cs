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

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Common
{
    /// <summary>
    /// Wrapper of the class BitConverter in order to support big endian
    /// Author: math
    /// </summary>
    internal class InternalBitConverter
    {

        private bool _IsLittleEndian = true;


        internal InternalBitConverter(bool isLittleEndian)
        {
            _IsLittleEndian = isLittleEndian;
        }


        internal UInt64 ToUInt64(byte[] value)
        {
            if (BitConverter.IsLittleEndian ^ _IsLittleEndian)
            {
                Array.Reverse(value);
            }
            return BitConverter.ToUInt64(value, 0);
        }


        internal UInt32 ToUInt32(byte[] value)
        {
            if (BitConverter.IsLittleEndian ^ _IsLittleEndian)
            {
                Array.Reverse(value);
            }
            return BitConverter.ToUInt32(value, 0);
        }


        internal ushort ToUInt16(byte[] value)
        {
            if (BitConverter.IsLittleEndian ^ _IsLittleEndian)
            {
                Array.Reverse(value);
            }
            return BitConverter.ToUInt16(value, 0);
        }


        internal string ToString(byte[] value)
        {
            if (BitConverter.IsLittleEndian ^ _IsLittleEndian)
            {
                Array.Reverse(value);
            }

            var enc = new UnicodeEncoding();            
            string result = enc.GetString(value);
            if (result.Contains("\0"))
            {
                result = result.Remove(result.IndexOf("\0"));
            }
            return result;
        }


        internal byte[] getBytes(ushort value)
        {
            var result = BitConverter.GetBytes(value);

            if (BitConverter.IsLittleEndian ^ _IsLittleEndian)
            {
                Array.Reverse(result);
            }
            return result;
        }


        internal byte[] getBytes(UInt32 value)
        {
            var result = BitConverter.GetBytes(value);

            if (BitConverter.IsLittleEndian ^ _IsLittleEndian)
            {
                Array.Reverse(result);
            }
            return result;
        }


        internal byte[] getBytes(UInt64 value)
        {
            var result = BitConverter.GetBytes(value);

            if (BitConverter.IsLittleEndian ^ _IsLittleEndian)
            {
                Array.Reverse(result);
            }
            return result;
        }

        internal List<byte> getBytes(List <UInt32> input)
        {
            var output = new List<byte>();

            foreach (var entry in input)
	        {
                output.AddRange(getBytes(entry));
            }
            return output;
        }
    }
}
