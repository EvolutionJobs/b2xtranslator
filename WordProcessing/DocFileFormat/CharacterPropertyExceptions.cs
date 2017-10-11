using b2xtranslator.CommonTranslatorLib;

namespace b2xtranslator.DocFileFormat
{

    public class CharacterPropertyExceptions : PropertyExceptions
    {
        /// <summary>
        /// Creates a CHPX wich doesn't modify anything.<br/>
        /// The grpprl list is empty
        /// </summary>
        public CharacterPropertyExceptions()
            : base()
        {

        }

        /// <summary>
        /// Parses the bytes to retrieve a CHPX
        /// </summary>
        /// <param name="bytes">The bytes starting with the istd</param>
        public CharacterPropertyExceptions(byte[] bytes) 
            : base(bytes)
        {

        }

        #region IVisitable Members

        public override void Convert<T>(T mapping)
        {
            ((IMapping<CharacterPropertyExceptions>)mapping).Apply(this);
        }

        #endregion
    }
}
