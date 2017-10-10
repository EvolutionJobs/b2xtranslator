

using System;
using System.Text;
using System.IO;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.Tools;

namespace b2xtranslator.PptFileFormat
{
    public class TextAtom : Record, ITextDataRecord
    {
        public string Text;

        public TextAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance, Encoding encoding)
            : base(_reader, size, typeCode, version, instance)
        {
            var bytes = new byte[size];
            this.Reader.Read(bytes, 0, (int)size);

            this.Text = new string(encoding.GetChars(bytes));
        }

        public override string ToString(uint depth)
        {
            return string.Format("{0}\n{1}Text = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1), Utils.StringInspect(this.Text));
        }

        #region ITextDataRecord Member

        private TextHeaderAtom _TextHeaderAtom;

        public TextHeaderAtom TextHeaderAtom
        {
            get
            {
                return this._TextHeaderAtom;
            }
            set
            {
                this._TextHeaderAtom = value;
            }
        }

        #endregion
    }

}
