using System;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class AutoNumberLevelDescriptor
    {
        /// <summary>
        /// Number format code
        /// </summary>
        public byte nfc;

        /// <summary>
        /// Offset into anld.rgxch that is the limit of the text that will be 
        /// displayed as the prefix of the auto number text
        /// </summary>
        public byte cxchTextBefore;

        /// <summary>
        /// anld.cxchTextBefore will be the beginning offset of the text in 
        /// the anld.rgxch that will be displayed as the suffix of an auto number. 
        /// The sum of anld.cxchTextBefore + anld.cxchTextAfter will be the limit 
        /// of the auto number suffix in anld.rgxch
        /// </summary>
        public byte cxchTextAfter;

        /// <summary>
        /// Justification code<br/>
        /// 0 left justify<br/>
        /// 1 center<br/>
        /// 2 right justify<br/>
        /// 3 left and right justfy
        /// </summary>
        public byte jc;

        /// <summary>
        /// When true, number generated will include previous levels
        /// </summary>
        public bool fPrev;

        /// <summary>
        /// When true, number will be displayed using hanging indent
        /// </summary>
        public bool fHang;

        /// <summary>
        /// When true, boldness of number will be determined by fBold
        /// </summary>
        public bool fSetBold;

        /// <summary>
        /// When true, italicness of number will be determined by fItalic
        /// </summary>
        public bool fSetItalic;

        /// <summary>
        /// When true, fSmallCaps will determine wheter number will be 
        /// displayed in small caps or not.
        /// </summary>
        public bool fSetSmallCaps;

        /// <summary>
        /// When true, fCaps will determine wheter number will be 
        /// displayed capitalized or not.
        /// </summary>
        public bool fSetCaps;

        /// <summary>
        /// When true, fStrike will determine wheter the number will be 
        /// displayed using strikethrough or not.
        /// </summary>
        public bool fSetStrike;

        /// <summary>
        /// When true, kul will determine the underlining state of 
        /// the auto number
        /// </summary>
        public bool fSetKul;

        /// <summary>
        /// When true, auto number will be displayed with a single 
        /// prefixing space character
        /// </summary>
        public bool fPrevSpace;

        /// <summary>
        /// Determines boldness of auto number when fSetBold is true
        /// </summary>
        public bool fBold;

        /// <summary>
        /// Determines italicness of auto number when fSetItalic is true
        /// </summary>
        public bool fItalic;

        /// <summary>
        /// Determines wheter auto number will be displayed using 
        /// small caps when fSetSmallCaps is true
        /// </summary>
        public bool fSmallCaps;

        /// <summary>
        /// Determines wheter auto number will be displayed using 
        /// caps when fSetCaps is true
        /// </summary>
        public bool fCaps;

        /// <summary>
        /// Determines wheter auto number will be displayed using 
        /// caps when fSetStrike is true
        /// </summary>
        public bool fStrike;

        /// <summary>
        /// Determines wheter auto number will be displayed with 
        /// underlining when fSetKul is true
        /// </summary>
        public byte kul;

        /// <summary>
        /// Color of auto number for Word 97.<br/>
        /// Unused in Word 2000
        /// </summary>
        public byte ico;

        /// <summary>
        /// Font code of auto number
        /// </summary>
        public short ftc;

        /// <summary>
        /// Font half point size (0 = auto)
        /// </summary>
        public ushort hps;

        /// <summary>
        /// Starting value
        /// </summary>
        public ushort iStartAt;

        /// <summary>
        /// Width of prefix text (same as indent)
        /// </summary>
        public ushort dxaIndent;

        /// <summary>
        /// Minimum space between number and paragraph
        /// </summary>
        public ushort dxaSpace;

        /// <summary>
        /// 24-bit color for Word 2000
        /// </summary>
        public int cv;

        /// <summary>
        /// Creates a new AutoNumberedListDataDescriptor with default values
        /// </summary>
        public AutoNumberLevelDescriptor()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a AutoNumberLevelDescriptor
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public AutoNumberLevelDescriptor(byte[] bytes)
        {
            if (bytes.Length == 20)
            {
                this.nfc = bytes[0];
                this.cxchTextBefore = bytes[1];
                this.cxchTextAfter = bytes[2];
                int b3 = (int)bytes[3];
                this.jc = Convert.ToByte(b3 & 0x03);
                this.fPrev = Utils.BitmaskToBool(b3, 0x04);
                this.fHang = Utils.BitmaskToBool(b3, 0x08);
                this.fSetBold = Utils.BitmaskToBool(b3, 0x10);
                this.fSetItalic = Utils.BitmaskToBool(b3, 0x20);
                this.fSetSmallCaps = Utils.BitmaskToBool(b3, 0x40);
                this.fSetCaps = Utils.BitmaskToBool(b3, 0x80);
                int b4 = (int)bytes[4];
                this.fSetStrike = Utils.BitmaskToBool(b4, 0x01);
                this.fSetKul = Utils.BitmaskToBool(b4, 0x02);
                this.fPrevSpace = Utils.BitmaskToBool(b4, 0x04);
                this.fBold = Utils.BitmaskToBool(b4, 0x08);
                this.fItalic = Utils.BitmaskToBool(b4, 0x10);
                this.fSmallCaps = Utils.BitmaskToBool(b4, 0x20);
                this.fCaps = Utils.BitmaskToBool(b4, 0x40);
                this.fStrike = Utils.BitmaskToBool(b4, 0x80);
                int b5 = (int)bytes[5];
                this.kul = (byte)(b5 & 0x07);
                this.ico = (byte)(b5 & 0xF1);
                this.ftc = System.BitConverter.ToInt16(bytes, 6);
                this.hps = System.BitConverter.ToUInt16(bytes, 8);
                this.iStartAt = System.BitConverter.ToUInt16(bytes, 10);
                this.dxaIndent = System.BitConverter.ToUInt16(bytes, 12);
                this.dxaSpace = System.BitConverter.ToUInt16(bytes, 14);
                this.cv = System.BitConverter.ToInt32(bytes, 16);
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct AutoNumberLevelDescriptor, the length of the struct doesn't match");
            }
        }

        private void setDefaultValues()
        {
 	        this.cv = 0;
            this.cxchTextAfter = 0;
            this.cxchTextBefore = 0;
            this.dxaIndent = 0;
            this.dxaSpace = 0;
            this.fBold = false;
            this.fCaps = false;
            this.fHang = false;
            this.fItalic = false;
            this.fPrev = false;
            this.fPrevSpace = false;
            this.fSetBold = false;
            this.fSetCaps = false;
            this.fSetItalic = false;
            this.fSetKul = false;
            this.fSetSmallCaps = false;
            this.fSetStrike = false;
            this.fSmallCaps = false;
            this.fStrike = false;
            this.ftc = 0;
            this.hps = 0;
            this.ico = 0;
            this.iStartAt = 0;
            this.jc = 0;
            this.kul = 0;
            this.nfc = 0;
        }
    }
}
