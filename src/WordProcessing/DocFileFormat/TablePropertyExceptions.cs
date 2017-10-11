using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class TablePropertyExceptions : PropertyExceptions
    {
        /// <summary>
        /// Creates a TAPX wich doesn't modify anything.<br/>
        /// The grpprl list is empty
        /// </summary>
        public TablePropertyExceptions()
            : base()
        {
        }

        /// <summary>
        /// Parses the bytes to retrieve a TAPX
        /// </summary>
        /// <param name="bytes">The bytes starting with the istd</param>
        public TablePropertyExceptions(byte[] bytes)
            : base(bytes)
        {
            //not yet implemented
        }

        /// <summary>
        /// Extracts the TAPX SPRMs out of a PAPX
        /// </summary>
        /// <param name="papx"></param>
        public TablePropertyExceptions(ParagraphPropertyExceptions papx, VirtualStream dataStream)
        {
            this.grpprl = new List<SinglePropertyModifier>();
            foreach (var sprm in papx.grpprl)
            {
                if (sprm.Type == SinglePropertyModifier.SprmType.TAP)
                {
                    this.grpprl.Add(sprm);
                }
                else if ((int)sprm.OpCode == 0x646b)
                {
                    IStreamReader reader = new VirtualStreamReader(dataStream);

                    //there is a native TAP in the data stream
                    uint fc = System.BitConverter.ToUInt32(sprm.Arguments, 0);
                    
                    //get the size of the following grpprl
                    //byte[] sizebytes = new byte[2];
                    //dataStream.Read(sizebytes, 2, (int)fc);
                    var sizebytes = reader.ReadBytes(fc, 2);
                    ushort grpprlSize = System.BitConverter.ToUInt16(sizebytes, 0);

                    //read the grpprl
                    //byte[] grpprlBytes = new byte[grpprlSize];
                    //dataStream.Read(grpprlBytes);
                    var grpprlBytes = reader.ReadBytes(grpprlSize);

                    //parse the grpprl
                    var externalPx = new PropertyExceptions(grpprlBytes);

                    foreach (var sprmExternal in externalPx.grpprl)
                    {
                        if (sprmExternal.Type == SinglePropertyModifier.SprmType.TAP)
                        {
                            this.grpprl.Add(sprmExternal);
                        }
                    }
                }
                
            }
        }

        #region IVisitable Members

        public override void Convert<T>(T mapping)
        {
            ((IMapping<TablePropertyExceptions>)mapping).Apply(this);
        }

        #endregion
    }
}
