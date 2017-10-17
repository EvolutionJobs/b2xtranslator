using System.Collections.Generic;
using System.Text;
using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    public class TextStyleAtom : Record, ITextDataRecord
    {
        public List<ParagraphRun> PRuns = new List<ParagraphRun>();
        public List<CharacterRun> CRuns = new List<CharacterRun>();

        public TextStyleAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        { }

        public override string ToString(uint depth)
        {
            var sb = new StringBuilder();
            sb.Append(base.ToString(depth));

            depth++;
            string indent = IndentationForDepth(depth);

            sb.AppendFormat("\n{0}Paragraph Runs:", indent);
            foreach (var pr in this.PRuns)
                sb.AppendFormat("\n{0}", pr.ToString(depth + 1));

            sb.AppendFormat("\n{0}Character Runs:", indent);
            foreach (var cr in this.CRuns)
                sb.AppendFormat("\n{0}", cr.ToString(depth + 1));

            return sb.ToString();
        }

        public virtual void AfterTextHeaderSet()
        { }

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
                this.AfterTextHeaderSet();
            }
        }

        #endregion
    }
}
