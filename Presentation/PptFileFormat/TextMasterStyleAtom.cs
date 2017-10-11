using b2xtranslator.OfficeDrawing;
using System.IO;
using b2xtranslator.Tools;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4003)]
    public class TextMasterStyleAtom : TextStyleAtom
    {
        public ushort IndentLevelCount;

        public byte[] Bytes;

        public TextMasterStyleAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Bytes = this.Reader.ReadBytes((int)this.BodySize);
            this.Reader.BaseStream.Position = 0;

            this.IndentLevelCount = this.Reader.ReadUInt16();

            for (int i = 0; i < this.IndentLevelCount; i++)
            {
                long pos = this.Reader.BaseStream.Position;

                if ((this.Instance >= 5)) // & (this.Instance < this.IndentLevelCount))
                {
                    ushort level = this.Reader.ReadUInt16(); 
                }                
             
                this.PRuns.Add(new ParagraphRun(this.Reader, true));

                TraceLogger.DebugInternal("Read paragraph run. Before pos = {0}, after pos = {1} of {2}: {3}",
                        pos, this.Reader.BaseStream.Position, this.Reader.BaseStream.Length,
                        this.PRuns[i].ToString());

                pos = this.Reader.BaseStream.Position;
                this.CRuns.Add(new CharacterRun(this.Reader));

                TraceLogger.DebugInternal("Read character run. Before pos = {0}, after pos = {1} of {2}: {3}",
                    pos, this.Reader.BaseStream.Position, this.Reader.BaseStream.Length,
                    this.CRuns[i].ToString());
            }

            //// XXX: I'm not sure why but in some cases there is trailing garbage -- flgr
            if (this.Reader.BaseStream.Position != this.Reader.BaseStream.Length)
            {
               this.Reader.BaseStream.Position = this.Reader.BaseStream.Length;
            }
        }

        public ParagraphRun ParagraphRunForIndentLevel(int level)
        {
            return this.PRuns[level];
        }

        public CharacterRun CharacterRunForIndentLevel(int level)
        {
            return this.CRuns[level];
        }
    }
}
