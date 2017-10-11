using System.Collections.Generic;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class AnnotationOwnerList : List<string>
    {
        public AnnotationOwnerList(FileInformationBlock fib, VirtualStream tableStream) : base()
        {
            tableStream.Seek(fib.fcGrpXstAtnOwners, System.IO.SeekOrigin.Begin);

            while (tableStream.Position < (fib.fcGrpXstAtnOwners + fib.lcbGrpXstAtnOwners))
            {
                this.Add(Utils.ReadXst(tableStream));
            }
        }
    }
}
