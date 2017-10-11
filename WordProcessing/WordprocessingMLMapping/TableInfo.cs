using b2xtranslator.DocFileFormat;
using b2xtranslator.Tools;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class TableInfo
    {
        /// <summary>
        /// 
        /// </summary>
        public bool fInTable;

        /// <summary>
        /// 
        /// </summary>
        public bool fTtp;

        /// <summary>
        /// 
        /// </summary>
        public bool fInnerTtp;

        /// <summary>
        /// 
        /// </summary>
        public bool fInnerTableCell;

        /// <summary>
        /// 
        /// </summary>
        public uint iTap;

        public TableInfo(ParagraphPropertyExceptions papx)
        {
            foreach (var sprm in papx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFInTable)
                {
                    this.fInTable = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFTtp)
                {
                    this.fTtp = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFInnerTableCell)
                {
                    this.fInnerTableCell = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPFInnerTtp)
                {
                    this.fInnerTtp = Utils.ByteToBool(sprm.Arguments[0]);
                }
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPItap)
                {
                    this.iTap = System.BitConverter.ToUInt32(sprm.Arguments, 0);
                    if (this.iTap > 0)
                        this.fInTable = true;
                }
                if ((int)sprm.OpCode == 0x66A)
                {
                    //add value!
                    this.iTap = System.BitConverter.ToUInt32(sprm.Arguments, 0);
                    if (this.iTap > 0)
                        this.fInTable = true;
                }
            }
        }
    }
}
