using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4013)]
    public class TextMasterStyle9Atom : Record
    {
        public List<ParagraphRun9> pruns = new List<ParagraphRun9>();
       
        public TextMasterStyle9Atom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            uint level = this.Reader.ReadUInt16();
            for (int i = 0; i < level; i++)
            {
                ParagraphMask pmask = (ParagraphMask)Reader.ReadUInt32();
                ParagraphRun9 pr = new ParagraphRun9();
                pr.mask = pmask;
                if ((pmask & ParagraphMask.BulletBlip) != 0)
                {
                    int bulletblipref = Reader.ReadInt16();
                    pr.bulletblipref = bulletblipref;
                }
                if ((pmask & ParagraphMask.BulletHasScheme) != 0)
                {
                    pr.fBulletHasAutoNumber = Reader.ReadInt16();
                }
                if ((pmask & ParagraphMask.BulletScheme) != 0)
                {
                    pr.bulletAutoNumberScheme = Reader.ReadUInt16();
                    pr.startAt = Reader.ReadInt16(); //start value
                }
                this.pruns.Add(pr);

                CharacterMask cmask = (CharacterMask)Reader.ReadUInt32();
                if ((cmask & CharacterMask.pp11ext) != 0)
                {
                    byte[] rest = Reader.ReadBytes(4);
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
        public Int16 startAt = -1;

        public bool BulletBlipReferencePresent
        {
            get { return (this.mask & ParagraphMask.BulletBlip) != 0; }
        }
    }
}
