
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgFunc : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgFunc;
        public ushort tab;


        public PtgFunc(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 3;
            this.Data = "";
            this.type = PtgType.Operator;
             this.tab = this.Reader.ReadUInt16();
            this.popSize = 1;
        }
    }
}
