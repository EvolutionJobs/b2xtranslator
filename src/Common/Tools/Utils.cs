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
using System.IO;
using System.Collections;
using System.Xml;

namespace DIaLOGIKa.b2xtranslator.Tools
{
    public class Utils
    {
        //Microsoft Office 2007 Beta 2 used namespaces that are not valid anymore.
        //This method should be used for all xml-code inside the binary Powerpoint file.
        //e.g. Themes, Layouts
        public static void replaceOutdatedNamespaces(ref XmlNode e)
        {
            string s = replaceOutdatedNamespaces(e.OuterXml);
            XmlDocument d2 = new XmlDocument();
            d2.LoadXml(s);
            e = d2.DocumentElement;
        }

        public static string replaceOutdatedNamespaces(string input)
        {
            string output = input.Replace("http://schemas.openxmlformats.org/drawingml/2006/3/main", "http://schemas.openxmlformats.org/drawingml/2006/main");
            output = output.Replace("http://schemas.openxmlformats.org/presentationml/2006/3/main", "http://schemas.openxmlformats.org/presentationml/2006/main");
            return output;
        }

        public static string ReadWString(Stream stream)
        {
            byte[] cch = new byte[1];
            stream.Read(cch, 0, cch.Length);

            byte[] chars = new byte[2 * cch[0]];
            stream.Read(chars, 0, chars.Length);

            return Encoding.Unicode.GetString(chars);
        }

        /// <summary>
        /// Read a length prefixed Unicode string from the given stream.
        /// The string must have the following structure:<br/>
        /// byte 1 - 4:         Character count (cch)<br/>
        /// byte 5 - (cch*2)+4: Unicode characters terminated by \0
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static string ReadLengthPrefixedUnicodeString(Stream stream)
        {
            byte[] cchBytes = new byte[4];
            stream.Read(cchBytes, 0, cchBytes.Length);
            Int32 cch = System.BitConverter.ToInt32(cchBytes, 0);

            //dont read the terminating zero
            byte[] stringBytes = new byte[cch*2];
            stream.Read(stringBytes, 0, stringBytes.Length);

            return Encoding.Unicode.GetString(stringBytes, 0, stringBytes.Length-2);
        }

        /// <summary>
        /// Read a length prefixed ANSI string from the given stream.
        /// The string must have the following structure:<br/>
        /// byte 1-4:       Character count (cch)<br/>
        /// byte 5-cch+4:   ANSI characters terminated by \0
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static string ReadLengthPrefixedAnsiString(Stream stream)
        {
            byte[] cchBytes = new byte[4];
            stream.Read(cchBytes, 0, cchBytes.Length);
            Int32 cch = System.BitConverter.ToInt32(cchBytes, 0);

            //dont read the terminating zero
            byte[] stringBytes = new byte[cch];
            stream.Read(stringBytes, 0, stringBytes.Length);

            if (cch > 0)
                return Encoding.ASCII.GetString(stringBytes, 0, stringBytes.Length - 1);
            else
                return null;
        }

        public static string ReadXstz(byte[] bytes, int pos)
        {
            byte[] xstz = new byte[System.BitConverter.ToUInt16(bytes, pos) * 2];
            Array.Copy(bytes, pos + 2, xstz, 0, xstz.Length);
            return Encoding.Unicode.GetString(xstz);
        }

        public static string ReadXst(Stream stream)
        {
            // read the char count
            byte[] cch = new byte[2];
            stream.Read(cch, 0, cch.Length);
            UInt16 charCount = System.BitConverter.ToUInt16(cch, 0);

            // read the string
            byte[] xst = new byte[charCount * 2];
            stream.Read(xst, 0, xst.Length);
            return Encoding.Unicode.GetString(xst);
        }

        public static string ReadXstz(Stream stream)
        {
            string xst = ReadXst(stream);
            
            //skip the termination
            byte[] termiantion = new byte[2];
            stream.Read(termiantion, 0, termiantion.Length);

            return xst;
        }

        public static string ReadShortXlUnicodeString(Stream stream)
        {
            byte[] cch = new byte[1];
            stream.Read(cch, 0, cch.Length);

            byte[] fHighByte = new byte[1];
            stream.Read(fHighByte, 0, fHighByte.Length);

            int rgbLength = cch[0];
            if (fHighByte[0] >= 0)
            {
                //double byte characters
                rgbLength *= 2;
            }

            byte[] rgb = new byte[rgbLength];
            stream.Read(rgb, 0, rgb.Length);

            if (fHighByte[0] >= 0)
            {
                return Encoding.Unicode.GetString(rgb);
            }
            else
            {
                Encoding enc = Encoding.GetEncoding(1252);
                return enc.GetString(rgb);
            }
        }

        public static int ArraySum(byte[] values)
        {
            int ret = 0;
            foreach (byte b in values)
            {
                ret += b;
            }
            return ret;
        }

        public static bool BitmaskToBool(Int32 value, Int32 mask)
        {
            return ((value & mask) == mask);
        }

        public static bool BitmaskToBool(UInt32 value, UInt32 mask)
        {
            return ((value & mask) == mask);
        }

        public static byte BitmaskToByte(UInt32 value, UInt32 mask)
        {
            value = value & mask;
            while ((mask & 0x1) != 0x1)
            {
                value = value >> 1;
                mask = mask >> 1;
            }
            return Convert.ToByte(value);
        }

        public static int BitmaskToInt(int value, int mask)
        {
            int ret = value & mask;
            //shift for all trailing zeros
            BitArray bits = new BitArray(new int[] { mask });
            foreach (bool bit in bits)
            {
                if (!bit)
                    ret = ret >> 1;
                else
                    break;
            }
            return ret;
        }

        public static Int32 BitmaskToInt32(Int32 value, Int32 mask)
        {
            value = value & mask;
            while ((mask & 0x1) != 0x1)
            {
                value = value >> 1;
                mask = mask >> 1;
            }
            return value;
        }

        public static UInt32 BitmaskToUInt32(UInt32 value, UInt32 mask)
        {
            value = value & mask;
            while ((mask & 0x1) != 0x1)
            {
                value = value >> 1;
                mask = mask >> 1;
            }
            return value;
        }

        public static UInt16 BitmaskToUInt16(UInt32 value, UInt32 mask)
        {
            return Convert.ToUInt16(BitmaskToUInt32(value, mask));
        }

        public static bool IntToBool(int value)
        {
            if (value == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        public static bool ByteToBool(byte value)
        {
            if (value == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static char[] ClearCharArray(char[] values)
        {
            char[] ret = new char[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                ret[i] = Convert.ToChar(0);
            }
            return ret;
        }

        public static int[] ClearIntArray(int[] values)
        {
            int[] ret = new int[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                ret[i] = 0;
            }
            return ret;
        }

        public static short[] ClearShortArray(ushort[] values)
        {
            short[] ret = new short[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                ret[i] = 0;
            }
            return ret;
        }

        public static UInt32 BitArrayToUInt32(BitArray bits)
        {
            double ret = 0;
            for (int i = 0; i < bits.Count; i++)
            {
                if (bits[i])
                {
                    ret += Math.Pow((double)2, (double)i);
                }
            }
            return (UInt32)ret;
        }

        public static BitArray BitArrayCopy(BitArray source, int sourceIndex, int copyCount)
        {
            bool[] ret = new bool[copyCount];

            int j = 0;
            for (int i = sourceIndex; i < (copyCount + sourceIndex); i++)
            {
                ret[j] = source[i];
                j++;
            }

            return new BitArray(ret);
        }

        public static string GetHashDump(byte[] bytes)
        {
            int colCount = 16;
            string ret = String.Format("({0:X04}) ", 0);

            int colCounter = 0;
            for (int i = 0; i < bytes.Length; i++)
            {
                if (colCounter == colCount)
                {
                    colCounter = 0;
                    ret += Environment.NewLine + String.Format("({0:X04}) ", i);
                }
                ret += String.Format("{0:X02} ", bytes[i]);
                colCounter++;
            }

            return ret;
        }

        // Would have been nice to use an extension method here... -- flgr
        public static String StringInspect(String s)
        {
            StringBuilder result = new StringBuilder("\"");

            foreach (char c in s)
            {
                switch (c)
                {
                    case '\r':
                        result.Append(@"\r");
                        break;

                    case '\n':
                        result.Append(@"\n");
                        break;

                    case '\v':
                        result.Append(@"\v");
                        break;

                    default:
                        if (Char.IsControl(c))
                            result.AppendFormat("\\x{0:X2}", (int)c);
                        else
                            result.Append(c);
                        break;
                }
            }

            result.Append("\"");

            return result.ToString();
        }

        public static String GetWritableString(String s)
        {
            StringBuilder result = new StringBuilder();
            foreach (char c in s)
            {
                if ((int)c >= 0x20)
                {
                    result.Append(c);
                }
            }
            return result.ToString();
        }
    }
}
