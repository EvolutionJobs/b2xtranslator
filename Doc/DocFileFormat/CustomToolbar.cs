using System.Collections.Generic;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class CustomToolbar : ByteStructure
    {
        /// <summary>
        /// Specifies the name of this custom toolbar.
        /// </summary>
        public string name;

        /// <summary>
        /// Signed integer that specifies the size, in bytes, of this structure excluding the name, cCtls, and rTBC fields. 
        /// Value is given by the following formula: cbTBData = sizeof(tb) + sizeof(rVisualData) + 12
        /// </summary>
        public int cbTBData;

        /// <summary>
        /// Structure of type TB, as specified in [MS-OSHARED], that contains toolbar data.
        /// </summary>
        public byte[] tb;

        public byte[] rVisualData;

        /// <summary>
        /// Signed integer that specifies the zero-based index of the Customization structure that 
        /// contains this structure in the rCustomizations array that contains the Customization 
        /// structure that contains this structure. <br/>
        /// Value MUST be greater or equal to 0x00000000 and MUST be less than the value of the 
        /// cCust field of the CTBWRAPPER structure that contains the rCustomizations array that 
        /// contains the Customization structure that contains this structure.
        /// </summary>
        public int iWCTB;

        /// <summary>
        /// Signed integer that specifies the number of toolbar controls in this toolbar.
        /// </summary>
        public int cCtls;

        /// <summary>
        /// Zero-based index array of TBC structures. <br/>
        /// The number of elements in this array MUST equal cCtls.
        /// </summary>
        public List<ToolbarControl> rTBC;

        public CustomToolbar(VirtualStreamReader reader)
            : base(reader, ByteStructure.VARIABLE_LENGTH)
        {
            this.name = Utils.ReadXst(reader.BaseStream);
            this.cbTBData = reader.ReadInt32();

            //cbTBData specifies the size of this structure excluding the name, cCtls, and rTBC fields
            //so it is the size of cbtb + tb + rVisualData + iWCTB + 4ignore bytes
            //so we can retrieve the size of tb:
            this.tb = reader.ReadBytes(this.cbTBData - 4 - 100 - 4 - 4);
            this.rVisualData = reader.ReadBytes(100);
            this.iWCTB = reader.ReadInt32();
            reader.ReadBytes(4);

            this.cCtls = reader.ReadInt32();
            this.rTBC = new List<ToolbarControl>();
            for (int i = 0; i < this.cCtls; i++)
            {   
                this.rTBC.Add(new ToolbarControl(reader));
            }
        }
    }
}
