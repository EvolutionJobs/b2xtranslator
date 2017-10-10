
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgFuncVar : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgFuncVar;
        public byte cparams; 
        public ushort tab; 
        public bool fCelFunc; 


        public PtgFuncVar(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 4;
            this.Data = "";
            this.type = PtgType.Operator;
            this.cparams = this.Reader.ReadByte();
            this.tab = this.Reader.ReadUInt16();

            this.fCelFunc = false;
            if ((this.tab & 0xF000) == 1)
            {
                this.fCelFunc = true; 
            }
            this.tab = (ushort)(this.tab & 0x7FFF); 
            
            this.popSize = (uint)(this.cparams);  
        }
    }
}
