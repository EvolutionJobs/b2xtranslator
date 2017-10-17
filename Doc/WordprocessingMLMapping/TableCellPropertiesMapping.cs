using System;
using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using System.Collections;
using b2xtranslator.Tools;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class TableCellPropertiesMapping : 
        PropertiesMapping,
        IMapping<TablePropertyExceptions>
    {
        private int _gridIndex;
        private int _cellIndex;
        private XmlElement _tcPr;
        private XmlElement _tcMar;
        private XmlElement _tcBorders;
        private List<short> _grid;
        private short[] _tGrid;

        private short _width;
        private Global.CellWidthType _ftsWidth;
        private TC80 _tcDef;

        private BorderCode _brcTop, _brcLeft, _brcRight, _brcBottom;

        /// <summary>
        /// The grind span of this cell
        /// </summary>
        private int _gridSpan;
        public int GridSpan
        {
            get { return this._gridSpan; }
        }

        private enum VerticalCellAlignment
        {
            top,
            center,
            bottom
        }

        public TableCellPropertiesMapping(XmlWriter writer, List<short> tableGrid, int gridIndex, int cellIndex)
            : base(writer)
        {
            this._tcPr = this._nodeFactory.CreateElement("w", "tcPr", OpenXmlNamespaces.WordprocessingML);
            this._tcMar = this._nodeFactory.CreateElement("w", "tcMar", OpenXmlNamespaces.WordprocessingML);
            this._tcBorders = this._nodeFactory.CreateElement("w", "tcBorders", OpenXmlNamespaces.WordprocessingML);
            this._gridIndex = gridIndex;
            this._grid = tableGrid;
            this._cellIndex = cellIndex;
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            //int lastBdr = getLastTabelBorderOccurrence(tapx.grpprl);

            for (int i=0; i< tapx.grpprl.Count; i++)
            {
                var sprm = tapx.grpprl[i];

                switch (sprm.OpCode)
	            {
                    //Table definition SPRM
                    case  SinglePropertyModifier.OperationCode.sprmTDefTable:
                        var tdef = new SprmTDefTable(sprm.Arguments);
                        this._tGrid = tdef.rgdxaCenter;
                        this._tcDef = tdef.rgTc80[this._cellIndex];

                        appendValueElement(this._tcPr, "textDirection", this._tcDef.textFlow.ToString(), false);

                        if (this._tcDef.vertMerge == Global.VerticalMergeFlag.fvmMerge)
                            appendValueElement(this._tcPr, "vMerge", "continue", false);
                        else if (this._tcDef.vertMerge == Global.VerticalMergeFlag.fvmRestart)
                            appendValueElement(this._tcPr, "vMerge", "restart", false);
                        else if (this._tcDef.vertMerge == Global.VerticalMergeFlag.fvmRestart)
                            appendValueElement(this._tcPr, "vMerge", "restart", false);

                        appendValueElement(this._tcPr, "vAlign", this._tcDef.vertAlign.ToString(), false);

                        if (this._tcDef.fFitText)
                            appendValueElement(this._tcPr, "tcFitText", "", false);

                        if (this._tcDef.fNoWrap)
                            appendValueElement(this._tcPr, "noWrap", "", true);

                        //_width = _tcDef.wWidth;
                        this._width = (short)(tdef.rgdxaCenter[this._cellIndex + 1] - tdef.rgdxaCenter[this._cellIndex]);
                        this._ftsWidth = this._tcDef.ftsWidth;

                        //borders
                        // if the sprm has a higher priority than the last sprmTTableBorder sprm in the list
                        //if (i > lastBdr)
                        //{
                        //    _brcTop = _tcDef.brcTop;
                        //    _brcLeft = _tcDef.brcLeft;
                        //    _brcRight = _tcDef.brcRight;
                        //    _brcBottom = _tcDef.brcBottom;
                        //}

                        break;

                    //margins
                    case SinglePropertyModifier.OperationCode.sprmTCellPadding:
                        byte first = sprm.Arguments[0];
                        byte lim = sprm.Arguments[1];
                        byte ftsMargin = sprm.Arguments[3];
                        short wMargin = System.BitConverter.ToInt16(sprm.Arguments, 4);

                        if (this._cellIndex >= first && this._cellIndex < lim)
                        {
                            var borderBits = new BitArray(new byte[] { sprm.Arguments[2] });
                            if (borderBits[0] == true)
                                appendDxaElement(this._tcMar, "top", wMargin.ToString(), true);
                            if (borderBits[1] == true)
                                appendDxaElement(this._tcMar, "left", wMargin.ToString(), true);
                            if (borderBits[2] == true)
                                appendDxaElement(this._tcMar, "bottom", wMargin.ToString(), true);
                            if (borderBits[3] == true)
                                appendDxaElement(this._tcMar, "right", wMargin.ToString(), true);
                        }
                        break;

                    //shading
                    case SinglePropertyModifier.OperationCode.sprmTDefTableShd:
                        //cell shading for cells 0-20
                        apppendCellShading(sprm.Arguments, this._cellIndex);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmTDefTableShd2nd:
                        //cell shading for cells 21-42
                        apppendCellShading(sprm.Arguments, this._cellIndex - 21);
                        break;
                    case SinglePropertyModifier.OperationCode.sprmTDefTableShd3rd:
                        //cell shading for cells 43-62
                        apppendCellShading(sprm.Arguments, this._cellIndex - 43);
                        break;

                    //width
                    case SinglePropertyModifier.OperationCode.sprmTCellWidth:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        if (this._cellIndex >= first && this._cellIndex < lim)
                        {
                            this._ftsWidth = (Global.CellWidthType)sprm.Arguments[2];
                            this._width = System.BitConverter.ToInt16(sprm.Arguments, 3);
                        }
                        break;

                    //vertical alignment
                    case SinglePropertyModifier.OperationCode.sprmTVertAlign:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        if (this._cellIndex >= first && this._cellIndex < lim)
                            appendValueElement(this._tcPr, "vAlign", ((VerticalCellAlignment)sprm.Arguments[2]).ToString(), true);
                        break;

                    //Autofit
                    case SinglePropertyModifier.OperationCode.sprmTFitText:
                        first = sprm.Arguments[0];
                        lim = sprm.Arguments[1];
                        if (this._cellIndex >= first && this._cellIndex < lim)
                            appendValueElement(this._tcPr, "tcFitText", sprm.Arguments[2].ToString(), true);
                        break;

                    //borders (cell definition)
                    case SinglePropertyModifier.OperationCode.sprmTSetBrc:
                        byte min = sprm.Arguments[0];
                        byte max = sprm.Arguments[1];
                        int bordersToApply = (int)sprm.Arguments[2] ;

                        if (this._cellIndex >= min && this._cellIndex < max)
                        {
                            var brcBytes = new byte[8];
                            Array.Copy(sprm.Arguments, 3, brcBytes, 0, 8);
                            var border = new BorderCode(brcBytes);
                            if(Utils.BitmaskToBool(bordersToApply, 0x01))
                            {
                                this._brcTop = border;
                            }
                            if(Utils.BitmaskToBool(bordersToApply, 0x02))
                            {
                                this._brcLeft = border;
                            }
                            if (Utils.BitmaskToBool(bordersToApply, 0x04))
                            {
                                this._brcBottom = border;
                            }
                            if (Utils.BitmaskToBool(bordersToApply, 0x08))
                            {
                                this._brcRight = border;
                            }
                        }
                        break;
                }
            }

            //width
            var tcW = this._nodeFactory.CreateElement("w", "tcW", OpenXmlNamespaces.WordprocessingML);
            var tcWType = this._nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
            var tcWVal = this._nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
            tcWType.Value = this._ftsWidth.ToString();
            tcWVal.Value = this._width.ToString();
            tcW.Attributes.Append(tcWType);
            tcW.Attributes.Append(tcWVal);
            this._tcPr.AppendChild(tcW);

            //grid span
            this._gridSpan = 1;
            if (this._width > this._grid[this._gridIndex])
            {
                //check the number of merged cells
                int w = this._grid[this._gridIndex];
                for (int i = this._gridIndex +1; i < this._grid.Count; i++)
                {
                    this._gridSpan++;
                    w += this._grid[i];
                    if (w >= this._width)
                        break;
                }

                appendValueElement(this._tcPr, "gridSpan", this._gridSpan.ToString(), true);
            }

            //append margins
            if (this._tcMar.ChildNodes.Count > 0)
            {
                this._tcPr.AppendChild(this._tcMar);
            }

            //append borders
            if (this._brcTop != null)
            {
                var topBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                appendBorderAttributes(this._brcTop, topBorder);
                addOrSetBorder(this._tcBorders, topBorder);
            }
            if (this._brcLeft != null)
            {
                var leftBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                appendBorderAttributes(this._brcLeft, leftBorder);
                addOrSetBorder(this._tcBorders, leftBorder);
            }
            if (this._brcBottom != null)
            {
                var bottomBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                appendBorderAttributes(this._brcBottom, bottomBorder);
                addOrSetBorder(this._tcBorders, bottomBorder);
            }
            if (this._brcRight != null)
            {
                var rightBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                appendBorderAttributes(this._brcRight, rightBorder);
                addOrSetBorder(this._tcBorders, rightBorder);
            }
            if (this._tcBorders.ChildNodes.Count > 0)
            {
                this._tcPr.AppendChild(this._tcBorders);
            }

            //write Properties
            if (this._tcPr.ChildNodes.Count > 0 || this._tcPr.Attributes.Count > 0)
            {
                this._tcPr.WriteTo(this._writer);
            }
        }

        private void apppendCellShading(byte[] sprmArg, int cellIndex)
        {
            //shading descriptor can have 10 bytes (Word 2000) or 2 bytes (Word 97)
            int shdLength = 2;
            if (sprmArg.Length % 10 == 0)
                shdLength = 10;

            var shdBytes = new byte[shdLength];

            //multiple cell can be formatted with the same SHD.
            //in this case there is only 1 SHD for all cells in the row.
            if ((cellIndex * shdLength) >= sprmArg.Length)
            {
                //use the first SHD
                cellIndex = 0;
            }

            Array.Copy(sprmArg, cellIndex * shdBytes.Length, shdBytes, 0, shdBytes.Length);
            
            var shd = new ShadingDescriptor(shdBytes);
            appendShading(this._tcPr, shd);
        }


        /// <summary>
        /// Returns the index of the last occurence of an sprmTTableBorders or sprmTTableBorders80 sprm.
        /// </summary>
        /// <param name="grpprl">The grpprl of sprms</param>
        /// <returns>The index or -1 if no sprm is in the list</returns>
        private int getLastTabelBorderOccurrence(List<SinglePropertyModifier> grpprl)
        {
            int index = -1;

            for (int i = 0; i < grpprl.Count; i++)
            {
                if (grpprl[i].OpCode == SinglePropertyModifier.OperationCode.sprmTTableBorders ||
                    grpprl[i].OpCode == SinglePropertyModifier.OperationCode.sprmTTableBorders80)
                {
                    index = i;
                }
            }

            return index;
        }
    }
}
