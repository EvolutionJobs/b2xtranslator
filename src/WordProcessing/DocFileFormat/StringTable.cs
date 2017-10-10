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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class StringTable : IVisitable
    {
        public bool fExtend;

        public int cData;

        public ushort cbExtra;

        public List<string> Strings;

        public List<ByteStructure> Data;

        //public List<ByteStructure> ExtraData;

        private Encoding _enc;

        public StringTable(Type dataType, VirtualStreamReader reader)
        {
            this.Strings = new List<string>();
            this.Data = new List<ByteStructure>();

            parse(dataType, reader, (uint)reader.BaseStream.Position);
        }

        public StringTable(Type dataType, VirtualStream tableStream, uint fc, uint lcb)
        {
            this.Strings = new List<string>();
            this.Data = new List<ByteStructure>();

            if (lcb > 0)
            {
                tableStream.Seek((long)fc, System.IO.SeekOrigin.Begin);
                parse(dataType, new VirtualStreamReader(tableStream), fc);
            }
        }

        private void parse(Type dataType, VirtualStreamReader reader, uint fc)
        {
            //read fExtend
            if (reader.ReadUInt16() == 0xFFFF)
            {
                //if the first 2 bytes are 0xFFFF the STTB contains unicode characters
                this.fExtend = true;
                this._enc = Encoding.Unicode;
            }
            else
            {
                //else the STTB contains 1byte characters and the fExtend field is non-existend
                //seek back to the beginning
                this.fExtend = false;
                this._enc = Encoding.ASCII;
                reader.BaseStream.Seek((long)fc, System.IO.SeekOrigin.Begin);
            }

            //read cData
            long cDataStart = reader.BaseStream.Position;
            var c = reader.ReadUInt16();
            if (c != 0xFFFF)
            {
                //cData is a 2byte unsigned Integer and the read bytes are already cData
                this.cData = (int)c;
            }
            else
            {
                //cData is a 4byte signed Integer, so we need to seek back
                reader.BaseStream.Seek((long)fc + cDataStart, System.IO.SeekOrigin.Begin);
                this.cData = reader.ReadInt32();
            }

            //read cbExtra
            this.cbExtra = reader.ReadUInt16();

            //read the strings and extra datas
            for (int i = 0; i < this.cData; i++)
            {
                int cchData = 0;
                int cbData = 0;
                if (this.fExtend)
                {
                    cchData = (int)reader.ReadUInt16();
                    cbData = cchData * 2;
                }
                else
                {
                    cchData = (int)reader.ReadByte();
                    cbData = cchData;
                }

                long posBeforeType = reader.BaseStream.Position;

                if (dataType == typeof(string))
                {
                    //It's a real string table
                    this.Strings.Add(this._enc.GetString(reader.ReadBytes(cbData)));
                }
                else
                {
                    //It's a modified string table that contains custom data
                    var constructor = dataType.GetConstructor(new Type[] { typeof(VirtualStreamReader), typeof(int) });
                    var data = (ByteStructure)constructor.Invoke(new object[] { reader, cbData });
                    this.Data.Add(data);
                }

                reader.BaseStream.Seek(posBeforeType + cbData, System.IO.SeekOrigin.Begin);

                //skip the extra byte
                reader.ReadBytes(this.cbExtra);
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<StringTable>)mapping).Apply(this);
        }

        #endregion
    }
}
