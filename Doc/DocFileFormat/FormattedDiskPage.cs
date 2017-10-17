using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class FormattedDiskPage 
    {
        public enum FKPType
        {
            Character,
            Paragraph
        }

        /// <summary>
        /// Specifies the type of the FKP
        /// </summary>
        public FKPType Type;

        /// <summary>
        /// The WordDocument stream
        /// </summary>
        public VirtualStream WordStream;

        /// <summary>
        /// Count of runs for that FKP
        /// </summary>
        public byte crun;

        /// <summary>
        /// Each value is the limit of a paragraph or run of exception text
        /// </summary>
        public int[] rgfc;

    }
}
