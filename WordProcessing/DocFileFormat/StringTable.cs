using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.StructuredStorage.Reader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace b2xtranslator.DocFileFormat
{
    public class StringTable :
        IVisitable
    {
        public bool fExtend;

        public int cData;

        public ushort cbExtra;

        public List<string> Strings;

        public List<ByteStructure> Data;

        Encoding _enc;

        public StringTable(Type dataType, VirtualStreamReader reader)
        {
            this.Strings = new List<string>();
            this.Data = new List<ByteStructure>();

            this.Parse(dataType, reader, (uint)reader.BaseStream.Position);
        }

        public StringTable(Type dataType, VirtualStream tableStream, uint fc, uint lcb)
        {
            this.Strings = new List<string>();
            this.Data = new List<ByteStructure>();

            if (lcb > 0)
            {
                tableStream.Seek((long)fc, SeekOrigin.Begin);
                this.Parse(dataType, new VirtualStreamReader(tableStream), fc);
            }
        }

        void Parse(Type dataType, VirtualStreamReader reader, uint fc)
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
                reader.BaseStream.Seek((long)fc, SeekOrigin.Begin);
            }

            //read cData
            long cDataStart = reader.BaseStream.Position;
            ushort c = reader.ReadUInt16();
            if (c != 0xFFFF)
            {
                //cData is a 2byte unsigned Integer and the read bytes are already cData
                this.cData = (int)c;
            }
            else
            {
                //cData is a 4byte signed Integer, so we need to seek back
                reader.BaseStream.Seek((long)fc + cDataStart, SeekOrigin.Begin);
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

                reader.BaseStream.Seek(posBeforeType + cbData, SeekOrigin.Begin);

                //skip the extra byte
                reader.ReadBytes(this.cbExtra);

                if (reader.BaseStream.Position == reader.BaseStream.Length)
                    break; // At EoF
            }
        }

        public void Convert<T>(T mapping) =>
            ((IMapping<StringTable>)mapping).Apply(this);
    }
}