using System.Collections.Generic;
using System.Text;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.CommonTranslatorLib;

namespace b2xtranslator.DocFileFormat
{
    public class CommandTable : IVisitable
    {
        public StringTable CommandStringTable;

        public List<MacroData> MacroDatas;

        public Dictionary<int, string> MacroNames;

        public List<KeyMapEntry> KeyMapEntries;

        public CustomToolbarWrapper CustomToolbars;

        private bool breakWhile;

        public CommandTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            tableStream.Seek(fib.fcCmds, System.IO.SeekOrigin.Begin);
            var reader = new VirtualStreamReader(tableStream);

            //byte[] bytes = reader.ReadBytes((int)fib.lcbCmds);
            this.MacroDatas = new List<MacroData>();
            this.KeyMapEntries = new List<KeyMapEntry>();
            this.MacroNames = new Dictionary<int, string>();

            //skip the version
            reader.ReadByte();

            //parse the commandtable
            while (reader.BaseStream.Position < (fib.fcCmds + fib.lcbCmds) && !this.breakWhile)
            {
                //read the type
                byte ch = reader.ReadByte();

                switch (ch)
                {
                    case 0x1:
                        //it's a PlfMcd
                        int iMacMcd = reader.ReadInt32();
                        for (int i = 0; i < iMacMcd; i++)
                        {
                            this.MacroDatas.Add(new MacroData(reader));
                        }
                        break;
                    case 0x2:
                        //it's a PlfAcd
                        //skip the ACDs
                        int iMacAcd = reader.ReadInt32();
                        reader.ReadBytes(iMacAcd * 4);
                        break;
                    case 0x3:
                        //Keymap Entries
                        int iMacKme = reader.ReadInt32();
                        for (int i = 0; i < iMacKme; i++)
                        {
                            this.KeyMapEntries.Add(new KeyMapEntry(reader));
                        }
                        break;
                    case 0x4:
                        //Keymap Entries
                        int iMacKmeInvalid = reader.ReadInt32();
                        for (int i = 0; i < iMacKmeInvalid; i++)
                        {
                            this.KeyMapEntries.Add(new KeyMapEntry(reader));
                        }
                        break;
                    case 0x10:
                        //it's a TcgSttbf
                        this.CommandStringTable = new StringTable(typeof(string), reader); 
                        break;
                    case 0x11:
                        //it's a MacroNames table
                        int iMacMn = reader.ReadInt16();
                        for (int i = 0; i < iMacMn; i++)
                        {
                            short ibst = reader.ReadInt16();
                            short cch = reader.ReadInt16();
                            this.MacroNames[ibst] = Encoding.Unicode.GetString(reader.ReadBytes(cch * 2));
                            //skip the terminating zero
                            reader.ReadBytes(2);
                        }
                        break;
                    case 0x12:
                        //it's a CTBWRAPPER structure
                        this.CustomToolbars = new CustomToolbarWrapper(reader);
                        break;
                    default:
                        this.breakWhile = true;
                        break;
                }
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<CommandTable>)mapping).Apply(this);
        }

        #endregion
    }
}
