using System;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class SprmTDefTable
    {
        public byte numberOfColumns;

        /// <summary>
        /// An array of 16-bit signed integer that specifies horizontal distance in twips. <br/>
        /// MUST be greater than or equal to -31680 and less than or equal to 31680.
        /// </summary>
        public short[] rgdxaCenter;

        /// <summary>
        /// An array of TC80 that specifies the default formatting for a cell in the table. <br/>
        /// Each TC80 in the array corresponds to the equivalent column in the table.<br/>
        /// If there are fewer TC80s than columns, the remaining columns are formatted with the default TC80 formatting. <br/>
        /// If there are more TC80s than columns, the excess TC80s MUST be ignored.
        /// </summary>
        public TC80[] rgTc80;

        public SprmTDefTable(byte[] bytes)
        {
            this.numberOfColumns = bytes[0];
            int pointer = 1;

            //read rgdxaCenter
            this.rgdxaCenter = new short[this.numberOfColumns + 1];
            for (int i = 0; i < this.numberOfColumns + 1 ; i++)
            {
                this.rgdxaCenter[i] = System.BitConverter.ToInt16(bytes, pointer);
                pointer += 2;
            }

            //read rgTc80
            this.rgTc80 = new TC80[this.numberOfColumns];
            for (int i = 0; i < this.numberOfColumns; i++)
            {
                var tc = new TC80();

                if (pointer < bytes.Length)
                {
                    //the flags
                    ushort flags = System.BitConverter.ToUInt16(bytes, pointer);
                    tc.horzMerge = (byte)Utils.BitmaskToInt((int)flags, 0x3);
                    tc.textFlow = (Global.TextFlow)Utils.BitmaskToInt((int)flags, 0x1C);
                    tc.vertMerge = (Global.VerticalMergeFlag)Utils.BitmaskToInt((int)flags, 0x60);
                    tc.vertAlign = (Global.VerticalAlign)Utils.BitmaskToInt((int)flags, 0x180);
                    tc.ftsWidth = (Global.CellWidthType)Utils.BitmaskToInt((int)flags, 0xE00);
                    tc.fFitText = Utils.BitmaskToBool(flags, 0x1000);
                    tc.fNoWrap = Utils.BitmaskToBool(flags, 0x2000);
                    tc.fHideMark = Utils.BitmaskToBool(flags, 0x4000);
                    pointer += 2;

                    //cell width
                    tc.wWidth = System.BitConverter.ToInt16(bytes, pointer);
                    pointer += 2;

                    //border top
                    var brcTopBytes = new byte[4];
                    Array.Copy(bytes, pointer, brcTopBytes, 0, 4);
                    tc.brcTop = new BorderCode(brcTopBytes);
                    pointer += 4;

                    //border left
                    var brcLeftBytes = new byte[4];
                    Array.Copy(bytes, pointer, brcLeftBytes, 0, 4);
                    tc.brcLeft = new BorderCode(brcLeftBytes);
                    pointer += 4;

                    //border bottom
                    var brcBottomBytes = new byte[4];
                    Array.Copy(bytes, pointer, brcBottomBytes, 0, 4);
                    tc.brcBottom = new BorderCode(brcBottomBytes);
                    pointer += 4;

                    //border top
                    var brcRightBytes = new byte[4];
                    Array.Copy(bytes, pointer, brcRightBytes, 0, 4);
                    tc.brcRight = new BorderCode(brcRightBytes);
                    pointer += 4;
                }

                this.rgTc80[i] = tc;
            }
        }
    }

    public struct TC80
    {
        /// <summary>
        /// A value from the following table that specifies how this cell merges horizontally with the neighboring cells in the same row. <br/>
        /// MUST be one of the following values:<br/>
        /// 0        The cell is not merged with the cells on either side of it.
        /// 1        The cell is one of a set of horizontally merged cells. It contributes its layout region to the set and its own contents are not rendered.
        /// 2, 3     The cell is the first cell in a set of horizontally merged cells. The contents and formatting of this cell extend into any consecutive cells following it that are designated as part of the merged set.
        /// </summary>
        public byte horzMerge;

        /// <summary>
        /// A value from the TextFlow enumeration that specifies rotation settings for the text in the cell.
        /// </summary>
        public Global.TextFlow textFlow;

        /// <summary>
        /// A value from the VerticalMergeFlag enumeration that specifies how this cell merges vertically with the cells above or below it.
        /// </summary>
        public Global.VerticalMergeFlag vertMerge;

        /// <summary>
        /// A value from the VerticalAlign enumeration that specifies how contents inside this cell are aligned.
        /// </summary>
        public Global.VerticalAlign vertAlign;

        /// <summary>
        /// An Fts that specifies the unit of measurement for the wWidth field in the TC80 structure.
        /// </summary>
        public Global.CellWidthType ftsWidth;

        /// <summary>
        /// Specifies whether the contents of the cell are to be stretched out such that the full cell width is used.
        /// </summary>
        public bool fFitText;

        /// <summary>
        /// When set, specifies that the preferred layout of the contents of this cell are as a single line, 
        /// and cell widths can be adjusted to accommodate long lines. <br/>
        /// This preference is ignored when the preferred width of this cell is set to ftsDxa.
        /// </summary>
        public bool fNoWrap;

        /// <summary>
        /// When set, specifies that this cell is rendered with no height if all cells in the row are empty.
        /// </summary>
        public bool fHideMark;

        /// <summary>
        /// An integer that specifies the preferred width of the cell. 
        /// The width includes cell margins, but does not include cell spacing. MUST be non-negative.<br/>
        /// The unit of measurement depends on ftsWidth.
        /// If ftsWidth is set to ftsPercent, the value is a fraction of the width of the entire table.
        /// </summary>
        public short wWidth;

        public BorderCode brcTop;

        public BorderCode brcLeft;

        public BorderCode brcBottom;

        public BorderCode brcRight;
    }
}
