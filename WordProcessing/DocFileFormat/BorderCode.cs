using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class BorderCode : IVisitable
    {
        public enum BorderType
        {
            none = 0,
            single,
            thick,
            Double,
            unused,
            hairline,
            dotted,
            dashed,
            dotDash,
            dotDotDash,
            triple,
            thinThickSmallGap,
            thickThinSmallGap,
            thinThickThinSmallGap,
            thinThickMediumGap,
            thickThinMediumGap,
            thinThickThinMediumGap,
            thinThickLargeGap,
            thickThinLargeGap,
            thinThickThinLargeGap,
            wave,
            doubleWave,
            dashSmallGap,
            dashDotStroked,
            threeDEmboss,
            threeDEngrave
        }

        /// <summary>
        /// 24-bit border color
        /// </summary>
        public int cv;

        /// <summary>
        /// Width of a single line in 1/8pt, max of 32pt
        /// </summary>
        public byte dptLineWidth;

        /// <summary>
        /// Border type code:
        /// 0 none
        /// 1 single
        /// 2 thick
        /// 3 double
        /// 5 hairline
        /// 6 dot
        /// 7 dash large gap
        /// 8 dot dash
        /// 9 dot dot dash
        /// 10 triple
        /// 11 thin-thick small gap
        /// 12 tick-thin small gap
        /// 13 thin-thick-thin small gap
        /// 14 thin-thick medium gap
        /// 15 thick-thin medium gap
        /// 16 thin-thick-thin medium gap
        /// 17 thin-thick large gap
        /// 18 thick-thin large gap
        /// 19 thin-thick-thin large gap
        /// 20 wave
        /// 21 double wave
        /// 22 dash small gap
        /// 23 dash dot stroked
        /// 24 emboss 3D
        /// 25 engrave 3D
        /// </summary>
        public byte brcType;

        /// <summary>
        /// The color of the Border.
        /// Unused if cv is set.
        /// </summary>
        public Global.ColorIdentifier ico;

        /// <summary>
        /// Width of space to maintain between border and text within border
        /// </summary>
        public int dptSpace;

        /// <summary>
        /// When true, border is drawn with shadow. Must be false when BRC is substructure of the TC
        /// </summary>
        public bool fShadow;

        /// <summary>
        /// When true, don't reverse the border
        /// </summary>
        public bool fFrame;

        /// <summary>
        /// It's a nil BRC, bytes are FFFF.
        /// </summary>
        public bool fNil;

        /// <summary>
        /// Creates a new BorderCode with default values
        /// </summary>
        public BorderCode()
        {
        }

        /// <summary>
        /// Parses the byte for a BRC
        /// </summary>
        /// <param name="bytes"></param>
        public BorderCode(byte[] bytes)
        {
            if (Utils.ArraySum(bytes) == bytes.Length * 255)
            {
                this.fNil = true;
            }
            else if (bytes.Length == 8)
            {
                //it's a border code of Word 2000/2003
                this.cv = System.BitConverter.ToInt32(bytes, 0);
                this.ico = Global.ColorIdentifier.auto;

                this.dptLineWidth = bytes[4];
                this.brcType = bytes[5];

                short val = System.BitConverter.ToInt16(bytes, 6);
                this.dptSpace = val & 0x001F;

                //not sure if this is correct, the values from the spec are definitly wrong:
                this.fShadow = Utils.BitmaskToBool(val, 0x20);
                this.fFrame = Utils.BitmaskToBool(val, 0x40);
            }
            else if (bytes.Length == 4)
            {
                //it's a border code of Word 97
                ushort val = System.BitConverter.ToUInt16(bytes, 0);
                this.dptLineWidth = (byte)(val & 0x00FF);
                this.brcType = (byte)((val & 0xFF00) >> 8);
                val = System.BitConverter.ToUInt16(bytes, 2);
                this.ico = (Global.ColorIdentifier)(val & 0x00FF);
                this.dptSpace = (val & 0x1F00) >> 8;
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct BRC, the length of the struct doesn't match");
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<BorderCode>)mapping).Apply(this);
        }

        #endregion
    }
}
