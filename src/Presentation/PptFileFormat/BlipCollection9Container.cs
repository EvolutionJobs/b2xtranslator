using System;
using System.Collections.Generic;
using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
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
                this.blip = (BitmapBlip)rec;
            } else if (rec is MetafilePictBlip) {
                this.mblip = (MetafilePictBlip)rec;
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
            while (this.Reader.BaseStream.Position < this.Reader.BaseStream.Length)
            {
                try
                {
                    var pr = new ParagraphRun9();
                    var pmask = (ParagraphMask)this.Reader.ReadUInt32();
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
                        pr.bulletAutoNumberScheme = this.Reader.ReadInt16();
                        pr.startAt = this.Reader.ReadInt16(); //start value
                    }
                    this.P9Runs.Add(pr);

                    var cmask = (CharacterMask)this.Reader.ReadUInt32();
                    if ((cmask & CharacterMask.pp11ext) != 0)
                    {
                        var rest = this.Reader.ReadBytes(4);
                    }

                    this.si = new TextSIException(this.Reader);
                }
                catch (Exception)
                {
                    //ignore
                }
                
            }
        }
    }
}
