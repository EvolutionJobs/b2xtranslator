using System.Collections.Generic;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class AnnotationReferenceExtraTable : List<AnnotationReferenceDescriptorExtra>
    {
        private const int ARTDPost10_LENGTH = 16;

        public AnnotationReferenceExtraTable(FileInformationBlock fib, VirtualStream tableStream)
        {
            if(fib.nFib >= FileInformationBlock.FibVersion.Fib2002)
            {
                tableStream.Seek((long)fib.fcAtrdExtra, System.IO.SeekOrigin.Begin);
                var reader = new VirtualStreamReader(tableStream);

                int n = (int)fib.lcbAtrdExtra / ARTDPost10_LENGTH;

                //read the n ATRDPost10 structs
                for (int i = 0; i < n; i++)
                {
                    this.Add(new AnnotationReferenceDescriptorExtra(reader, ARTDPost10_LENGTH));        
                }
            }
        }
    }
}
