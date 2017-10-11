using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using System.Collections.Generic;

namespace b2xtranslator.DocFileFormat
{
    public sealed class StwStructure : Dictionary<string, string>
    {
        public StwStructure(VirtualStream tableStream, uint fc, uint lcb) : base()
        {
            tableStream.Seek((long)fc, System.IO.SeekOrigin.Begin);

            // parse the names
            var names = new StringTable(typeof(string), new VirtualStreamReader(tableStream));

            // parse the values
            var values = new List<string>();
            while (tableStream.Position < fc+lcb)
                values.Add(Utils.ReadXst(tableStream));

            // map to the dictionary
            if (names.Strings.Count == values.Count)
                for (int i = 0; i < names.Strings.Count; i++)
                    this.Add(names.Strings[i], values[i]);
        }
    }
}
