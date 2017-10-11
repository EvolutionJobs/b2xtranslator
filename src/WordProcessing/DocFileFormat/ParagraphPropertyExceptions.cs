using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class ParagraphPropertyExceptions : PropertyExceptions
    {
        /// <summary>
        /// Index to style descriptor of the style from which the 
        /// paragraph inherits its paragraph and character properties
        /// </summary>
        public ushort istd;

        /// <summary>
        /// Creates a PAPX wich doesn't modify anything.<br/>
        /// The grpprl list is empty
        /// </summary>
        public ParagraphPropertyExceptions() : base()
        {
        }

        /// <summary>
        /// Parses the bytes to retrieve a PAPX
        /// </summary>
        /// <param name="bytes">The bytes starting with the istd</param>
        public ParagraphPropertyExceptions(byte[] bytes, VirtualStream dataStream)
            : base(new List<byte>(bytes).GetRange(2, bytes.Length-2).ToArray())
        {
            if (bytes.Length != 0)
            {
                this.istd = System.BitConverter.ToUInt16(bytes, 0);
            }

            //There is a SPRM that points to an offset in the data stream, 
            //where a list of SPRM is saved.
            foreach (var sprm in this.grpprl)
            {
                if(sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPHugePapx || (int)sprm.OpCode == 0x6646)
                {
                    IStreamReader reader = new VirtualStreamReader(dataStream);
                    uint fc = System.BitConverter.ToUInt32(sprm.Arguments, 0);

                    //parse the size of the external grpprl
                    var sizebytes = new byte[2];
                    dataStream.Read(sizebytes, 0, 2, (int)fc);
                    ushort size = System.BitConverter.ToUInt16(sizebytes, 0);
                    
                    //parse the external grpprl
                    //byte[] grpprlBytes = new byte[size];
                    //dataStream.Read(grpprlBytes);
                    var grpprlBytes = reader.ReadBytes(size);
                    var externalPx = new PropertyExceptions(grpprlBytes);

                    //assign the external grpprl
                    this.grpprl = externalPx.grpprl;

                    //remove the sprmPHugePapx
                    this.grpprl.Remove(sprm);
                }
            }
        }

        #region IVisitable Members

        public override void Convert<T>(T mapping)
        {
            ((IMapping<ParagraphPropertyExceptions>)mapping).Apply(this);
        }

        #endregion
    }
}
