

using System;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    public enum TextType
    {
        Title = 0,
        Body,
        Notes,
        Outline,
        Other,
        CenterBody,
        CenterTitle,
        HalfBody,
        QuarterBody
    };

    [OfficeRecord(3999)]
    public class TextHeaderAtom : Record
    {
        public TextType TextType;

        public TextAtom TextAtom;

        public TextStyleAtom TextStyleAtom;

        public TextHeaderAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.TextType = (TextType) this.Reader.ReadUInt32();
        }

        public void HandleTextDataRecord(ITextDataRecord tdRecord)
        {
            tdRecord.TextHeaderAtom = this;

            var textAtom = tdRecord as TextAtom;
            var tsAtom = tdRecord as TextStyleAtom;

            if (textAtom != null)
            {
                this.TextAtom = textAtom;
            }
            else if (tsAtom != null)
            {
                this.TextStyleAtom = tsAtom;
            }
            else
            {
                throw new NotImplementedException("Unhandled text data record type " + tdRecord.GetType().ToString());
            }
        }
    }

}
