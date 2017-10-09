using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class CommentAuthorTable : List<string>
    {
        public CommentAuthorTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            tableStream.Seek(fib.fcGrpXstAtnOwners, System.IO.SeekOrigin.Begin);
            VirtualStreamReader reader = new VirtualStreamReader(tableStream);

            while (tableStream.Position < (fib.fcGrpXstAtnOwners + fib.lcbGrpXstAtnOwners))
            {
                Int16 cch = reader.ReadInt16();
                this.Add(Encoding.Unicode.GetString(reader.ReadBytes(cch * 2)));
            }
        }
    }
}
