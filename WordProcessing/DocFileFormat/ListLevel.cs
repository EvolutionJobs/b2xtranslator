using System.Text;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class ListLevel : ByteStructure
    {
        public enum FollowingChar
        {
            tab = 0,
            space,
            nothing
        }

        /// <summary>
        /// Start at value for this list level
        /// </summary>
        public int iStartAt;

        /// <summary>
        /// Number format code (see anld.nfc for a list of options)
        /// </summary>
        public byte nfc;

        /// <summary>
        /// Alignment (left, right, or centered) of the paragraph number.
        /// </summary>
        public byte jc;

        /// <summary>
        /// True if the level turns all inherited numbers to arabic, 
        /// false if it preserves their number format code (nfc)
        /// </summary>
        public bool fLegal;

        /// <summary>
        /// True if the level‘s number sequence is not restarted by 
        /// higher (more significant) levels in the list
        /// </summary>
        public bool fNoRestart;

        /// <summary>
        /// Word 6.0 compatibility option: equivalent to anld.fPrev (see ANLD)
        /// </summary>
        public bool fPrev;

        /// <summary>
        /// Word 6.0 compatibility option: equivalent to anld.fPrevSpace (see ANLD)
        /// </summary>
        public bool fPrevSpace;

        /// <summary>
        /// True if this level was from a converted Word 6.0 document. <br/>
        /// If it is true, all of the Word 6.0 compatibility options become 
        /// valid otherwise they are ignored.
        /// </summary>
        public bool fWord6;

        /// <summary>
        /// Contains the character offsets into the LVL’s XST of the inherited numbers of previous levels. <br/>
        /// The XST contains place holders for any paragraph numbers contained in the text of the number, 
        /// and the place holder contains the ilvl of the inherited number, 
        /// so lvl.xst[lvl.rgbxchNums[0]] == the level of the first inherited number in this level.
        /// </summary>
        public byte[] rgbxchNums;

        /// <summary>
        /// The type of character following the number text for the paragraph.
        /// </summary>
        public FollowingChar ixchFollow;

        /// <summary>
        /// Word 6.0 compatibility option: equivalent to anld.dxaSpace (see ANLD). <br/>
        /// For newer versions indent to remove if we remove this numbering.
        /// </summary>
        public int dxaSpace;

        /// <summary>
        /// Word 6.0 compatibility option: equivalent to anld.dxaIndent (see ANLD).<br/>
        /// Unused in newer versions.
        /// </summary>
        public int dxaIndent;

        /// <summary>
        /// Length, in bytes, of the LVL‘s grpprlChpx.
        /// </summary>
        public byte cbGrpprlChpx;

        /// <summary>
        /// Length, in bytes, of the LVL‘s grpprlPapx.
        /// </summary>
        public byte cbGrpprlPapx;

        /// <summary>
        /// Limit of levels that we restart after.
        /// </summary>
        public byte ilvlRestartLim;

        /// <summary>
        /// A grfhic that specifies HTML incompatibilities of the level.
        /// </summary>
        public byte grfhic;

        /// <summary>
        /// 
        /// </summary>
        public ParagraphPropertyExceptions grpprlPapx;

        /// <summary>
        /// 
        /// </summary>
        public CharacterPropertyExceptions grpprlChpx;

        /// <summary>
        /// 
        /// </summary>
        public string xst;

        /// <summary>
        /// Parses the given StreamReader to retrieve a LVL struct
        /// </summary>
        /// <param name="bytes"></param>
        public ListLevel(VirtualStreamReader reader, int length)
            : base(reader, length)
        {
            long startPos = this._reader.BaseStream.Position;

            //parse the fix part
            this.iStartAt = this._reader.ReadInt32();
            this.nfc = this._reader.ReadByte();
            int flag = this._reader.ReadByte();
            this.jc = (byte)(flag & 0x03);
            this.fLegal = Utils.BitmaskToBool(flag, 0x04);
            this.fNoRestart = Utils.BitmaskToBool(flag, 0x08);
            this.fPrev = Utils.BitmaskToBool(flag, 0x10);
            this.fPrevSpace = Utils.BitmaskToBool(flag, 0x20);
            this.fWord6 = Utils.BitmaskToBool(flag, 0x40);
            this.rgbxchNums = new byte[9];
            for (int i = 0; i < 9; i++)
            {
                this.rgbxchNums[i] = this._reader.ReadByte();
            }
            this.ixchFollow = (FollowingChar)this._reader.ReadByte();

            this.dxaSpace = this._reader.ReadInt32();
            this.dxaIndent = this._reader.ReadInt32();

            this.cbGrpprlChpx = this._reader.ReadByte();
            this.cbGrpprlPapx = this._reader.ReadByte();

            this.ilvlRestartLim = this._reader.ReadByte();
            this.grfhic = this._reader.ReadByte();
            
            //parse the variable part

            //read the group of papx sprms
            //this papx has no istd, so use PX to parse it
            var px = new PropertyExceptions(this._reader.ReadBytes(this.cbGrpprlPapx));
            this.grpprlPapx = new ParagraphPropertyExceptions
            {
                grpprl = px.grpprl
            };

            //read the group of chpx sprms
            this.grpprlChpx = new CharacterPropertyExceptions(this._reader.ReadBytes(this.cbGrpprlChpx));

            //read the number text
            short strLen = this._reader.ReadInt16();
            this.xst = Encoding.Unicode.GetString(this._reader.ReadBytes(strLen * 2));

            long endPos = this._reader.BaseStream.Position;
            this._reader.BaseStream.Seek(startPos, System.IO.SeekOrigin.Begin);
            this._rawBytes = this._reader.ReadBytes((int)(endPos - startPos));
        }
    }
}
