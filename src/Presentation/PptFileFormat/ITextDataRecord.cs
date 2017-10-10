


namespace b2xtranslator.PptFileFormat
{
    /// <summary>
    /// Marker interface used for marking Records that contain text data
    /// that is associated with a TextHeaderAtom record.
    /// </summary>
    public interface ITextDataRecord
    {
        /// <summary>
        /// TextHeaderAtom corresponding to this record.
        /// </summary>
        TextHeaderAtom TextHeaderAtom
        {
            get;
            set;
        }
    }
}
