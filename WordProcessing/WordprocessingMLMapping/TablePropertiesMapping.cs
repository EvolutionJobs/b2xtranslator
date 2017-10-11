using System;
using System.Collections.Generic;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Tools;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class TablePropertiesMapping :
        PropertiesMapping,
        IMapping<TablePropertyExceptions>
    {
        private XmlElement _tblPr;
        private XmlElement _tblGrid;
        private XmlElement _tblBorders;
        private StyleSheet _styles;
        private List<short> _grid;
        private BorderCode brcLeft, brcTop, brcBottom, brcRight, brcHorz, brcVert;

        private enum WidthType
        {
            nil,
            auto,
            pct,
            dxa
        }

        public TablePropertiesMapping(XmlWriter writer, StyleSheet styles, List<short> grid)
            : base(writer)
        {
            this._styles = styles;
            this._tblPr = this._nodeFactory.CreateElement("w", "tblPr", OpenXmlNamespaces.WordprocessingML);
            this._tblBorders = this._nodeFactory.CreateElement("w", "tblBorders", OpenXmlNamespaces.WordprocessingML);
            this._grid = grid;
        }

        public void Apply(TablePropertyExceptions tapx)
        {
            var tblBorders = this._nodeFactory.CreateElement("w", "tblBorders", OpenXmlNamespaces.WordprocessingML);
            var tblCellMar = this._nodeFactory.CreateElement("w", "tblCellMar", OpenXmlNamespaces.WordprocessingML);
            var tblLayout = this._nodeFactory.CreateElement("w", "tblLayout", OpenXmlNamespaces.WordprocessingML);
            var tblpPr = this._nodeFactory.CreateElement("w", "tblpPr", OpenXmlNamespaces.WordprocessingML);
            var layoutType = this._nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
            layoutType.Value = "fixed";
            short tblIndent = 0;
            short gabHalf = 0;
            short marginLeft = 0;
            short marginRight = 0;

            foreach (var sprm in tapx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    case SinglePropertyModifier.OperationCode.sprmTDxaGapHalf:
                        gabHalf = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        break;

                    //table definition
                    case SinglePropertyModifier.OperationCode.sprmTDefTable:
                        var tDef = new SprmTDefTable(sprm.Arguments);
                        //Workaround for retrieving the indent of the table:
                        //In some files there is a indent but no sprmTWidthIndent is set.
                        //For this cases we can calculate the indent of the table by getting the 
                        //first boundary of the TDef and adding the padding of the cells
                        tblIndent = System.BitConverter.ToInt16(sprm.Arguments, 1);
                        //add the gabHalf
                        tblIndent += gabHalf;
                        //If there follows a real sprmTWidthIndent, this value will be overwritten
                        break;

                    //preferred table width
                    case SinglePropertyModifier.OperationCode.sprmTTableWidth:
                        var fts = (WidthType)sprm.Arguments[0];
                        short width = System.BitConverter.ToInt16(sprm.Arguments, 1);
                        var tblW = this._nodeFactory.CreateElement("w", "tblW", OpenXmlNamespaces.WordprocessingML);
                        var w = this._nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                        w.Value = width.ToString();
                        var type = this._nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
                        type.Value = fts.ToString();
                        tblW.Attributes.Append(type);
                        tblW.Attributes.Append(w);
                        this._tblPr.AppendChild(tblW);
                        break;

                    //justification
                    case SinglePropertyModifier.OperationCode.sprmTJc:
                    case  SinglePropertyModifier.OperationCode.sprmTJcRow:
                        appendValueElement(this._tblPr, "jc", ((Global.JustificationCode)sprm.Arguments[0]).ToString(), true);
                        break;

                    //indent
                    case SinglePropertyModifier.OperationCode.sprmTWidthIndent:
                        tblIndent = System.BitConverter.ToInt16(sprm.Arguments, 1);
                        break;

                    //style
                    case SinglePropertyModifier.OperationCode.sprmTIstd:
                    case SinglePropertyModifier.OperationCode.sprmTIstdPermute:
                        short styleIndex = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        if(this._styles.Styles.Count> styleIndex)
                        {
                            string id = StyleSheetMapping.MakeStyleId(this._styles.Styles[styleIndex]);
                            if(id != "TableNormal")
                            {
                                appendValueElement(this._tblPr, "tblStyle", id, true);
                            }
                        }
                        break;

                    //bidi
                    case SinglePropertyModifier.OperationCode.sprmTFBiDi:
                    case SinglePropertyModifier.OperationCode.sprmTFBiDi90:
                        appendValueElement(this._tblPr, "bidiVisual", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString(), true);
                        break;

                    //table look
                    case SinglePropertyModifier.OperationCode.sprmTTlp:
                        appendValueElement(this._tblPr, "tblLook", string.Format("{0:x4}", System.BitConverter.ToInt16(sprm.Arguments, 2)), true);
                        break;

                    //autofit
                    case SinglePropertyModifier.OperationCode.sprmTFAutofit:
                        if (sprm.Arguments[0] == 1)
                            layoutType.Value = "auto";
                        break;

                    //default cell padding (margin)
                    case SinglePropertyModifier.OperationCode.sprmTCellPadding:
                    case SinglePropertyModifier.OperationCode.sprmTCellPaddingDefault:
                    case SinglePropertyModifier.OperationCode.sprmTCellPaddingOuter:
                        byte grfbrc = sprm.Arguments[2];
                        short wMar = System.BitConverter.ToInt16(sprm.Arguments, 4);
                        if (Utils.BitmaskToBool((int)grfbrc, 0x01))
                            appendDxaElement(tblCellMar, "top", wMar.ToString(), true);
                        if (Utils.BitmaskToBool((int)grfbrc, 0x02))
                            marginLeft = wMar;
                        if (Utils.BitmaskToBool((int)grfbrc, 0x04))
                            appendDxaElement(tblCellMar, "bottom", wMar.ToString(), true);
                        if (Utils.BitmaskToBool((int)grfbrc, 0x08))
                            marginRight = wMar;
                        break;

                    //row count
                    case SinglePropertyModifier.OperationCode.sprmTCHorzBands:
                        appendValueElement(this._tblPr, "tblStyleRowBandSize", sprm.Arguments[0].ToString(), true);
                        break;

                    //col count
                    case SinglePropertyModifier.OperationCode.sprmTCVertBands:
                        appendValueElement(this._tblPr, "tblStyleColBandSize", sprm.Arguments[0].ToString(), true);
                        break;

                    //overlap
                    case SinglePropertyModifier.OperationCode.sprmTFNoAllowOverlap:
                        bool noOverlap = Utils.ByteToBool(sprm.Arguments[0]);
                        string tblOverlapVal = "overlap";
                        if (noOverlap)
                            tblOverlapVal = "never";
                        appendValueElement(this._tblPr, "tblOverlap", tblOverlapVal, true);
                        break;

                    //shading
                    case SinglePropertyModifier.OperationCode.sprmTSetShdTable:
                        var desc = new ShadingDescriptor(sprm.Arguments);
                        appendShading(this._tblPr, desc);
                        break;

                    //borders 80 exceptions
                    case SinglePropertyModifier.OperationCode.sprmTTableBorders80:
                        var brc80 = new byte[4];
                        //top border
                        Array.Copy(sprm.Arguments, 0, brc80, 0, 4);
                        this.brcTop = new BorderCode(brc80);
                        //left
                        Array.Copy(sprm.Arguments, 4, brc80, 0, 4);
                        this.brcLeft = new BorderCode(brc80);
                        //bottom
                        Array.Copy(sprm.Arguments, 8, brc80, 0, 4);
                        this.brcBottom = new BorderCode(brc80);
                        //right
                        Array.Copy(sprm.Arguments, 12, brc80, 0, 4);
                        this.brcRight = new BorderCode(brc80);
                        //inside H
                        Array.Copy(sprm.Arguments, 16, brc80, 0, 4);
                        this.brcHorz = new BorderCode(brc80);
                        //inside V
                        Array.Copy(sprm.Arguments, 20, brc80, 0, 4);
                        this.brcVert = new BorderCode(brc80);
                        break;

                    //border exceptions
                    case SinglePropertyModifier.OperationCode.sprmTTableBorders:
                        var brc = new byte[8];
                        //top border
                        Array.Copy(sprm.Arguments, 0, brc, 0, 8);
                        this.brcTop = new BorderCode(brc);
                        //left
                        Array.Copy(sprm.Arguments, 8, brc, 0, 8);
                        this.brcLeft = new BorderCode(brc);
                        //bottom
                        Array.Copy(sprm.Arguments, 16, brc, 0, 8);
                        this.brcBottom = new BorderCode(brc);
                        //right
                        Array.Copy(sprm.Arguments, 24, brc, 0, 8);
                        this.brcRight = new BorderCode(brc);
                        //inside H
                        Array.Copy(sprm.Arguments, 32, brc, 0, 8);
                        this.brcHorz = new BorderCode(brc);
                        //inside V
                        Array.Copy(sprm.Arguments, 40, brc, 0, 8);
                        this.brcVert = new BorderCode(brc);
                        break;

                    //floating table properties
                    case SinglePropertyModifier.OperationCode.sprmTPc:
                        byte flag = sprm.Arguments[0];
                        var pcVert = (Global.VerticalPositionCode)((flag & 0x30) >> 4);
                        var pcHorz = (Global.HorizontalPositionCode)((flag & 0xC0) >> 6);
                        appendValueAttribute(tblpPr, "horzAnchor", pcHorz.ToString());
                        appendValueAttribute(tblpPr, "vertAnchor", pcVert.ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmTDxaFromText:
                        appendValueAttribute(tblpPr, "leftFromText", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmTDxaFromTextRight:
                        appendValueAttribute(tblpPr, "rightFromText", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmTDyaFromText:
                        appendValueAttribute(tblpPr, "topFromText", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmTDyaFromTextBottom:
                        appendValueAttribute(tblpPr, "bottomFromText", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmTDxaAbs:
                        appendValueAttribute(tblpPr, "tblpX", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                    case SinglePropertyModifier.OperationCode.sprmTDyaAbs:
                        appendValueAttribute(tblpPr, "tblpY", System.BitConverter.ToInt16(sprm.Arguments, 0).ToString());
                        break;
                }
            }

            //indent
            if (tblIndent != 0)
            {
                var tblInd = this._nodeFactory.CreateElement("w", "tblInd", OpenXmlNamespaces.WordprocessingML);
                var tblIndW = this._nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                tblIndW.Value = tblIndent.ToString();
                tblInd.Attributes.Append(tblIndW);
                var tblIndType = this._nodeFactory.CreateAttribute("w", "type", OpenXmlNamespaces.WordprocessingML);
                tblIndType.Value = "dxa";
                tblInd.Attributes.Append(tblIndType);
                this._tblPr.AppendChild(tblInd);
            }

            //append floating props
            if (tblpPr.Attributes.Count > 0)
            {
                this._tblPr.AppendChild(tblpPr);
            }

            //set borders
            if (this.brcTop != null)
            {
                var topBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "top", OpenXmlNamespaces.WordprocessingML);
                appendBorderAttributes(this.brcTop, topBorder);
                addOrSetBorder(this._tblBorders, topBorder);
            }
            if (this.brcLeft != null)
            {
                var leftBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "left", OpenXmlNamespaces.WordprocessingML);
                appendBorderAttributes(this.brcLeft, leftBorder);
                addOrSetBorder(this._tblBorders, leftBorder);
            }
            if (this.brcBottom != null)
            {
                var bottomBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "bottom", OpenXmlNamespaces.WordprocessingML);
                appendBorderAttributes(this.brcBottom, bottomBorder);
                addOrSetBorder(this._tblBorders, bottomBorder);
            }
            if (this.brcRight != null)
            {
                var rightBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "right", OpenXmlNamespaces.WordprocessingML);
                appendBorderAttributes(this.brcRight, rightBorder);
                addOrSetBorder(this._tblBorders, rightBorder);
            }
            if (this.brcHorz != null)
            {
                var insideHBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideH", OpenXmlNamespaces.WordprocessingML);
                appendBorderAttributes(this.brcHorz, insideHBorder);
                addOrSetBorder(this._tblBorders, insideHBorder);
            }
            if (this.brcVert != null)
            {
                var insideVBorder = this._nodeFactory.CreateNode(XmlNodeType.Element, "w", "insideV", OpenXmlNamespaces.WordprocessingML);
                appendBorderAttributes(this.brcVert, insideVBorder);
                addOrSetBorder(this._tblBorders, insideVBorder);
            }
            if (this._tblBorders.ChildNodes.Count > 0)
            {
                this._tblPr.AppendChild(this._tblBorders);
            }

            //append layout type
            tblLayout.Attributes.Append(layoutType);
            this._tblPr.AppendChild(tblLayout);

            //append margins
            if (marginLeft == 0 && gabHalf != 0)
            {
                appendDxaElement(tblCellMar, "left", gabHalf.ToString(), true);
            }
            else
            {
                appendDxaElement(tblCellMar, "left", marginLeft.ToString(), true);
            }
            if (marginRight == 0 && gabHalf != 0)
            {
                appendDxaElement(tblCellMar, "right", gabHalf.ToString(), true);
            }
            else
            {
                appendDxaElement(tblCellMar, "right", marginRight.ToString(), true);
            }
            this._tblPr.AppendChild(tblCellMar);

            //write Properties
            if (this._tblPr.ChildNodes.Count > 0 || this._tblPr.Attributes.Count > 0)
            {
                this._tblPr.WriteTo(this._writer);
            }

            //append the grid
            this._tblGrid = this._nodeFactory.CreateElement("w", "tblGrid", OpenXmlNamespaces.WordprocessingML);
            foreach (short colW in this._grid)
            {
                var gridCol = this._nodeFactory.CreateElement("w", "gridCol", OpenXmlNamespaces.WordprocessingML);
                var gridColW = this._nodeFactory.CreateAttribute("w", "w", OpenXmlNamespaces.WordprocessingML);
                gridColW.Value = colW.ToString();
                gridCol.Attributes.Append(gridColW);
                this._tblGrid.AppendChild(gridCol);
            }
            this._tblGrid.WriteTo(this._writer);
        }
    }
}
