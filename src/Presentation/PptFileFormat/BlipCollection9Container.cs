using System;
using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecord(2040)]
    public class BlipCollection9Container : RegularContainer
    {
        public BlipCollection9Container(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(2041)]
    public class BlipEntityAtom : Record
    {
        public BitmapBlip blip;
        public MetafilePictBlip mblip;

        public BlipEntityAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

            byte winBlipType = this.Reader.ReadByte();
            byte unused = this.Reader.ReadByte();

            var rec = Record.ReadRecord(this.Reader);

            if (rec is BitmapBlip)
            {
                blip = (BitmapBlip)rec;
            } else if (rec is MetafilePictBlip) {
                mblip = (MetafilePictBlip)rec;
            }
        }
    }

    [OfficeRecord(4012)]
    public class StyleTextProp9Atom : Record
    {
        public List<ParagraphRun9> P9Runs = new List<ParagraphRun9>();
        public TextSIException si;
        
        public StyleTextProp9Atom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            while (Reader.BaseStream.Position < Reader.BaseStream.Length)
            {
                try
                {
                    var pr = new ParagraphRun9();
                    var pmask = (ParagraphMask)Reader.ReadUInt32();
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
                        pr.bulletAutoNumberScheme = Reader.ReadInt16();
                        pr.startAt = Reader.ReadInt16(); //start value
                    }
                    P9Runs.Add(pr);

                    var cmask = (CharacterMask)Reader.ReadUInt32();
                    if ((cmask & CharacterMask.pp11ext) != 0)
                    {
                        var rest = Reader.ReadBytes(4);
                    }

                    si = new TextSIException(Reader);
                }
                catch (Exception)
                {
                    //ignore
                }
                
            }
        }
    }
}
