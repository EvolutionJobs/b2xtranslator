

using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4001)]
    public class TextRunStyleAtom : TextStyleAtom
    {
        public TextRunStyleAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        { }

        /// <summary>
        /// Can only read data after text header has been set. So automatic verification doesn't work.
        /// </summary>
        override public bool DoAutomaticVerifyReadToEnd
        {
            get { return false; }
        }

        override public void AfterTextHeaderSet()
        {
            var textAtom = this.TextHeaderAtom.TextAtom;

            /* This can legitimately happen... */
            if (textAtom == null)
            {
                this.Reader.BaseStream.Position = this.Reader.BaseStream.Length;
                return;
            }

            uint seenLength = 0;
            while (seenLength < textAtom.Text.Length + 1)
            {
                long pos = this.Reader.BaseStream.Position;
                uint length = this.Reader.ReadUInt32();

                var run = new ParagraphRun(this.Reader, false);
                run.Length = length;
                this.PRuns.Add(run);

                /*TraceLogger.DebugInternal("Read paragraph run. Before pos = {0}, after pos = {1} of {2}: {3}",
                    pos, this.Reader.BaseStream.Position, this.Reader.BaseStream.Length,
                    run);*/

                seenLength += length;
            }

            //TraceLogger.DebugInternal();

            seenLength = 0;
            while (seenLength < textAtom.Text.Length + 1)
            {
                uint length = this.Reader.ReadUInt32();

                var run = new CharacterRun(this.Reader);
                run.Length = length;
                this.CRuns.Add(run);

                seenLength += length;
            }

            this.VerifyReadToEnd();
        }
    }
}
