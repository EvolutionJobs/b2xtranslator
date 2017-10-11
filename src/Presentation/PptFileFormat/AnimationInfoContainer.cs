using System;
using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4116)]
    public class AnimationInfoContainer : RegularContainer
    {
        public AnimationInfoContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(4081)]
    public class AnimationInfoAtom : Record
    {
        public byte[] dimColor;
        public short flags;
        public byte[] soundIdRef;
        public int delayTime;
        public short orderID;
        public ushort slideCount;
        public AnimBuildTypeEnum animBuildType;
        public byte animEffect;
        public byte animEffectDirection;
        public AnimAfterEffectEnum animAfterEffect;
        public TextBuildSubEffectEnum textBuildSubEffect;
        public byte oleVerb;

        public bool fReverse;
        public bool fAutomatic;
        public bool fSound;
        public bool fStopSound;
        public bool fPlay;
        public bool fSynchronous;
        public bool fHide;
        public bool fAnimateBg;

        public AnimationInfoAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

            this.dimColor = this.Reader.ReadBytes(4);
            this.flags = this.Reader.ReadInt16();

            this.fReverse = Tools.Utils.BitmaskToBool(this.flags, 0x1 << 0);
            this.fAutomatic = Tools.Utils.BitmaskToBool(this.flags, 0x1 << 2);
            this.fSound = Tools.Utils.BitmaskToBool(this.flags, 0x1 << 4);
            this.fStopSound = Tools.Utils.BitmaskToBool(this.flags, 0x1 << 6);
            this.fPlay = Tools.Utils.BitmaskToBool(this.flags, 0x1 << 8);
            this.fSynchronous = Tools.Utils.BitmaskToBool(this.flags, 0x1 << 10);
            this.fHide = Tools.Utils.BitmaskToBool(this.flags, 0x1 << 12);
            this.fAnimateBg = Tools.Utils.BitmaskToBool(this.flags, 0x1 << 14);

            short reserved = this.Reader.ReadInt16();
            this.soundIdRef = this.Reader.ReadBytes(4);
            this.delayTime = this.Reader.ReadInt32();
            this.orderID = this.Reader.ReadInt16();
            this.slideCount = this.Reader.ReadUInt16();
            this.animBuildType = (AnimBuildTypeEnum)this.Reader.ReadByte();
            this.animEffect = this.Reader.ReadByte();
            this.animEffectDirection = this.Reader.ReadByte();
            this.animAfterEffect = (AnimAfterEffectEnum)this.Reader.ReadByte();
            this.textBuildSubEffect = (TextBuildSubEffectEnum)this.Reader.ReadByte();
            this.oleVerb = this.Reader.ReadByte();

            if (this.Reader.BaseStream.Position != this.Reader.BaseStream.Length)
            {
                this.Reader.BaseStream.Position = this.Reader.BaseStream.Length;
            }
        }

        [FlagsAttribute]
        public enum AnimationFlagsMask : uint
        {
            None = 0,

            fReverse = 3 << 14,
            fAutomatic = 3 << 12,
            fSound = 3 << 10,
            fStopSound = 3 << 8,
            fPlay = 3 << 6,
            fSynchronous = 3 << 4,
            fHide = 3 << 2,
            fAnimateBg = 3
        }

        public enum AnimBuildTypeEnum: byte
        {
            FollowMaster = 0xFE,

            NoBuild = 0x00,
            OneBuild = 0x01,
            Level1Build = 0x02,
            Level2Build = 0x03,
            Level3Build = 0x04,
            Level4Build = 0x05,
            Level5Build = 0x06,
            GraphBySeries = 0x07,
            GraphByCategory = 0x08,
            GraphByElementInSeries = 0x09,
            GraphByElementInCategory = 0x0A
        }

        public enum AnimAfterEffectEnum : byte
        {
            NoAfterEffect = 0x00,
            Dim = 0x01,
            Hide = 0x02,
            HideImmediately = 0x03
        }

        public enum TextBuildSubEffectEnum : byte
        {
            BuildByNone = 0x00,
            BuildByWord = 0x01,
            BuildByCharacter = 0x02
        }
    }

}
