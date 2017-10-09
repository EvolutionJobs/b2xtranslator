using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class StwStructure : Dictionary<String, String>
    {
        public StwStructure(VirtualStream tableStream, UInt32 fc, UInt32 lcb) : base()
        {
            tableStream.Seek((long)fc, System.IO.SeekOrigin.Begin);

            // parse the names
            StringTable names = new StringTable(typeof(String), new VirtualStreamReader(tableStream));

            // parse the values
            List<string> values = new List<string>();
            while (tableStream.Position < fc+lcb)
            {
                values.Add(Utils.ReadXst(tableStream));
            }

            // map to the dictionary
            if (names.Strings.Count == values.Count)
            {
                for (int i = 0; i < names.Strings.Count; i++)
                {
                    this.Add(names.Strings[i], values[i]);
                }
            }
        }
    }
}
