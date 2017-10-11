using System.Collections.Generic;
using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4013)]
    public class TextMasterStyle9Atom : Record
    {
        public List<ParagraphRun9> pruns = new List<ParagraphRun9>();
       
        public TextMasterStyle9Atom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            uint level = this.Reader.ReadUInt16();
            for (int i = 0; i < level; i++)
            {
                var pmask = (ParagraphMask)this.Reader.ReadUInt32();
                var pr = new ParagraphRun9();
                pr.mask = pmask;
                if ((pmask & ParagraphMask.BulletBlip) != 0)
                {
                    int bulletblipref = this.Reader.ReadInt16();
                    pr.bulletblipref = bulletblipref;
                }
                if ((pmask & ParagraphMask.BulletHasScheme) != 0)
                {
                    pr.fBulletHasAutoNumber = this.Reader.ReadInt16();
                }
                if ((pmask & ParagraphMask.BulletScheme) != 0)
                {
                    pr.bulletAutoNumberScheme = this.Reader.ReadUInt16();
                    pr.startAt = this.Reader.ReadInt16(); //start value
                }
                this.pruns.Add(pr);

                var cmask = (CharacterMask)this.Reader.ReadUInt32();
                if ((cmask & CharacterMask.pp11ext) != 0)
                {
                    var rest = this.Reader.ReadBytes(4);
                }

            }
        }
    }

    public class ParagraphRun9
    {
        public ParagraphMask mask;
        public int bulletblipref;
        public int fBulletHasAutoNumber;
        public int bulletAutoNumberScheme = -1;
        public short startAt = -1;

        public bool BulletBlipReferencePresent
        {
            get { return (this.mask & ParagraphMask.BulletBlip) != 0; }
        }
    }
}
