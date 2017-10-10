using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class StyleSheetInformation
    {
        public struct LatentStyleData
        {
            public uint grflsd;
            public bool fLocked;
        }

        /// <summary>
        /// Count of styles in stylesheet
        /// </summary>
        public ushort cstd;
	
        /// <summary>
        /// Length of STD Base as stored in a file
        /// </summary>
        public ushort cbSTDBaseInFile;
	
        /// <summary>
        /// Are built-in stylenames stored?
        /// </summary>
        public bool fStdStylenamesWritten;
						
        /// <summary>
        /// Max sti known when this file was written
        /// </summary>
        public ushort stiMaxWhenSaved;

        /// <summary>
        /// How many fixed-index istds are there?
        /// </summary>
	    public ushort istdMaxFixedWhenSaved;

        /// <summary>
        /// Current version of built-in stylenames
        /// </summary>
	    public ushort nVerBuiltInNamesWhenSaved;

        /// <summary>
        /// This is a list of the default fonts for this style sheet.<br/>
        /// The first is for ASCII characters (0-127), the second is for East Asian characters, 
        /// and the third is the default font for non-East Asian, non-ASCII text.
        /// </summary>
	    public ushort[] rgftcStandardChpStsh;	

	    /// <summary>
	    /// Size of each lsd in mpstilsd<br/>
        /// The count of lsd's is stiMaxWhenSaved
	    /// </summary>
        public ushort cbLSD;

        /// <summary>
        /// latent style data (size == stiMaxWhenSaved upon save!)
        /// </summary>
	    public LatentStyleData[] mpstilsd;	

        /// <summary>
        /// Parses the bytes to retrieve a StyleSheetInformation
        /// </summary>
        /// <param name="bytes"></param>
        public StyleSheetInformation(byte[] bytes)
        {
            this.cstd = System.BitConverter.ToUInt16(bytes, 0);
            this.cbSTDBaseInFile = System.BitConverter.ToUInt16(bytes, 2);
            if(bytes[4] == 1)
            {
                this.fStdStylenamesWritten = true;
            }
            //byte 5 is spare
            this.stiMaxWhenSaved = System.BitConverter.ToUInt16(bytes, 6);
            this.istdMaxFixedWhenSaved = System.BitConverter.ToUInt16(bytes, 8);
            this.nVerBuiltInNamesWhenSaved = System.BitConverter.ToUInt16(bytes, 10);

            this.rgftcStandardChpStsh = new ushort[4];
            this.rgftcStandardChpStsh[0] = System.BitConverter.ToUInt16(bytes, 12);
            this.rgftcStandardChpStsh[1] = System.BitConverter.ToUInt16(bytes, 14);
            this.rgftcStandardChpStsh[2] = System.BitConverter.ToUInt16(bytes, 16);
            if (bytes.Length > 18)
            {
                this.rgftcStandardChpStsh[3] = System.BitConverter.ToUInt16(bytes, 18);
            }

            //not all stylesheet contain latent styles
            if (bytes.Length > 20)
            {
                this.cbLSD = System.BitConverter.ToUInt16(bytes, 20);
                this.mpstilsd = new LatentStyleData[this.stiMaxWhenSaved];
                for (int i = 0; i < this.mpstilsd.Length; i++)
                {
                    var lsd = new LatentStyleData
                    {
                        grflsd = System.BitConverter.ToUInt32(bytes, 22 + (i * this.cbLSD))
                    };
                    lsd.fLocked = Utils.BitmaskToBool((int)lsd.grflsd, 0x1);
                    this.mpstilsd[i] = lsd;
                }
            }
        }
    }
}
