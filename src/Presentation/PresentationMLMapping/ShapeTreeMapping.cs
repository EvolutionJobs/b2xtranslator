

using System;
using System.Collections.Generic;
using System.Text;
using b2xtranslator.PptFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OfficeDrawing;
using System.Reflection;
using b2xtranslator.OpenXmlLib;
using System.Drawing;
using b2xtranslator.Tools;
using System.Collections;
using System.IO;

namespace b2xtranslator.PresentationMLMapping
{
    public class ShapeTreeMapping :
        AbstractOpenXmlMapping,
        IMapping<PPDrawing>
    {
        protected int _idCnt;
        protected ConversionContext _ctx;
        protected string _footertext;
        protected string _headertext;
        protected string _datetext;
        
        public PresentationMapping<RegularContainer> parentSlideMapping = null;
        public Dictionary<AnimationInfoContainer, int> animinfos = new Dictionary<AnimationInfoContainer, int>();
        public StyleTextProp9Atom ShapeStyleTextProp9Atom = null;

        private SortedDictionary<int, int> ColumnWidthsByYPos;
        private List<int> RowHeights;
        //private List<ShapeContainer> Lines;
        private SortedList<int, SortedList<int, ShapeContainer>> verticallinelist;
        private SortedList<int, SortedList<int, ShapeContainer>> horizontallinelist;

        public ShapeTreeMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            this._ctx = ctx;
        }

        public void DynamicApply(Record record)
        {
            // Call Apply(record) with dynamic dispatch (selection based on run-time type of record)
            var method = this.GetType().GetMethod("Apply", new Type[] { record.GetType() });

            //TraceLogger.DebugInternal(method.ToString());

            try
            {
                method.Invoke(this, new object[] { record });
            }
            catch (TargetInvocationException e)
            {
                TraceLogger.DebugInternal(e.InnerException.ToString());
                throw e.InnerException;
            }
        }

        public void Apply(PPDrawing drawing)
        {
            Apply((RegularContainer) drawing);
            writeVML();
        }

        public void Apply(DrawingContainer drawingContainer)
        {
            var group = drawingContainer.FirstChildWithType<GroupContainer>();
            IEnumerator<Record> iter = group.Children.GetEnumerator();
            iter.MoveNext();

            var header = iter.Current as ShapeContainer;
            WriteGroupShapeProperties(header);

            while (iter.MoveNext())
                DynamicApply(iter.Current);
        }

        //used to give each group a unique identifyer
        private int groupcounter = -10;
        public void Apply(GroupContainer group)
        {
            var gsr = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<GroupShapeRecord>();
            var anchor = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<ClientAnchor>();
            var chanchor = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<ChildAnchor>();

            foreach (var ops in group.FirstChildWithType<ShapeContainer>().AllChildrenWithType<ShapeOptions>())
            {
                if (ops.OptionsByID.ContainsKey(ShapeOptions.PropertyId.tableProperties))
                {
                    uint TABLEFLAGS = ops.OptionsByID[ShapeOptions.PropertyId.tableProperties].op;
                    if (Tools.Utils.BitmaskToBool(TABLEFLAGS, 0x1))
                    {
                        //this group is a table
                        ApplyTable(group, TABLEFLAGS);
                        return;
                    }
                }
            }

            this._writer.WriteStartElement("p", "grpSp", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "nvGrpSpPr", OpenXmlNamespaces.PresentationML);

            int id = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<Shape>().spid;
            WriteCNvPr(id, "");
            //WriteCNvPr(--groupcounter, "");
            this._writer.WriteElementString("p", "cNvGrpSpPr", OpenXmlNamespaces.PresentationML, "");
            this._writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");
            this._writer.WriteEndElement(); //nvGrpSpPr

            this._writer.WriteStartElement("p", "grpSpPr", OpenXmlNamespaces.PresentationML);

            if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
            {
                this._writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);

                this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                this._writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                this._writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteStartElement("a", "chOff", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("x", gsr.rcgBounds.Left.ToString());
                this._writer.WriteAttributeString("y", gsr.rcgBounds.Top.ToString());
                this._writer.WriteEndElement();

                this._writer.WriteStartElement("a", "chExt", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("cx", (gsr.rcgBounds.Right - gsr.rcgBounds.Left).ToString());
                this._writer.WriteAttributeString("cy", (gsr.rcgBounds.Bottom - gsr.rcgBounds.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteEndElement();
            }
            else if (chanchor != null && chanchor.Right >= chanchor.Left && chanchor.Bottom >= chanchor.Top)
            {
                this._writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);

                this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("x", (chanchor.Left).ToString());
                this._writer.WriteAttributeString("y", (chanchor.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("cx", (chanchor.Right - chanchor.Left).ToString());
                this._writer.WriteAttributeString("cy", (chanchor.Bottom - chanchor.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteStartElement("a", "chOff", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("x", gsr.rcgBounds.Left.ToString());
                this._writer.WriteAttributeString("y", gsr.rcgBounds.Top.ToString());
                this._writer.WriteEndElement();

                this._writer.WriteStartElement("a", "chExt", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("cx", (gsr.rcgBounds.Right - gsr.rcgBounds.Left).ToString());
                this._writer.WriteAttributeString("cy", (gsr.rcgBounds.Bottom - gsr.rcgBounds.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteEndElement();
            }

            this._writer.WriteEndElement(); //grpSpPr

            bool first = true;
            foreach (var record in group.Children)
            {
                //ignore first
                if (first)
                {
                    first = false;
                }
                else
                {
                    DynamicApply(record);
                }
            }

            this._writer.WriteEndElement(); //grpSp
        }

        private int GetGridSpanCount(ChildAnchor anch, int col)
        {
            int count = 0;
            int availableWidth = anch.rcgBounds.Width;
            int currentLeft = anch.Left;

            while (availableWidth > 0)
            {
                availableWidth -= this.ColumnWidthsByYPos[currentLeft];
                currentLeft += this.ColumnWidthsByYPos[currentLeft];
                count++;
            }            

            return count;
        }

        private int GetRowSpanCount(ChildAnchor anch, int row)
        {
            try
            {
                int count = 0;
                int availableHeight = anch.rcgBounds.Height;

                while (availableHeight > 0 && this.RowHeights.Count > row + count)
                {
                    availableHeight -= this.RowHeights[row + count];
                    if (availableHeight >= 0) count++; else return 0;
                }

                return count;
            }
            catch (Exception)
            {                
                throw;
            }
           
        }

        public void ApplyTable(GroupContainer group, uint TABLEFLAGS)
        {
            var gsr = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<GroupShapeRecord>();
            var sh = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<Shape>();
            var anchor = group.FirstChildWithType<ShapeContainer>().FirstChildWithType<ClientAnchor>();

            this.RowHeights = new List<int>();
            foreach (var ops in group.FirstChildWithType<ShapeContainer>().AllChildrenWithType<ShapeOptions>())
            {
                if (ops.OptionsByID.ContainsKey(ShapeOptions.PropertyId.tableRowProperties))
                {
                    uint TableRowPropertiesCount = ops.OptionsByID[ShapeOptions.PropertyId.tableRowProperties].op;
                    var data = ops.OptionsByID[ShapeOptions.PropertyId.tableRowProperties].opComplex;
                    ushort nElems = BitConverter.ToUInt16(data, 0);
                    ushort nElemsAlloc = BitConverter.ToUInt16(data, 2);
                    ushort cbElem = BitConverter.ToUInt16(data, 4);
                    for (int i = 0; i < nElems; i++)
                    {
                        int height = BitConverter.ToInt32(data, 6 + i * cbElem);
                        //this is a workaround for a bug
                        //it should be analysed when to use 0 height
                        if (height > 53)
                        {
                            this.RowHeights.Add(height);
                        }
                        else
                        {
                            this.RowHeights.Add(0);
                        }
                    }
                }                
            }

            var Cells = new List<ShapeContainer>();
            //Lines = new List<ShapeContainer>();
            var LinesByPosition = new Dictionary<string, ShapeContainer>();
           
            var tablelist = new SortedList<int,SortedList<int,ShapeContainer>>();
            this.verticallinelist = new SortedList<int, SortedList<int, ShapeContainer>>();
            this.horizontallinelist = new SortedList<int, SortedList<int, ShapeContainer>>();


            foreach (var scontainer in group.AllChildrenWithType<ShapeContainer>())
            {
                var anch = scontainer.FirstChildWithType<ChildAnchor>();
                var cd = scontainer.FirstChildWithType<ClientData>();

                foreach (var shape in scontainer.AllChildrenWithType<Shape>())
                {
                    if (Utils.getPrstForShape(shape.Instance) == "rect")
                    {
                        if (!tablelist.ContainsKey(anch.Top)) tablelist.Add(anch.Top, new SortedList<int, ShapeContainer>());
                        tablelist[anch.Top].Add(anch.Left, scontainer);
                    }
                    else if (Utils.getPrstForShape(shape.Instance) == "line")
                    {
                        if (anch.Top == anch.Bottom)
                        {
                            //horizontal
                            if (!this.horizontallinelist.ContainsKey(anch.Top)) this.horizontallinelist.Add(anch.Top, new SortedList<int, ShapeContainer>());
                            this.horizontallinelist[anch.Top].Add(anch.Left, scontainer);
                        }
                        else
                        {
                            //vertical
                            if (!this.verticallinelist.ContainsKey(anch.Top)) this.verticallinelist.Add(anch.Top, new SortedList<int, ShapeContainer>());
                            this.verticallinelist[anch.Top].Add(anch.Left, scontainer);
                        }

                        //Lines.Add(scontainer);
                    }
                    else
                    {
                        string s = Utils.getPrstForShape(shape.Instance);
                    }
                } 
            }

            this.ColumnWidthsByYPos = new SortedDictionary<int, int>();
            var ColumnIndices = new Dictionary<int, int>(); //this list will contain all column limits
            foreach (var rowlist in tablelist.Values)
            {
                //rowlist contains all cells in a row
                foreach (int y in rowlist.Keys)
                {
                    int w = rowlist[y].FirstChildWithType<ChildAnchor>().rcgBounds.Width;
                    if (!this.ColumnWidthsByYPos.ContainsKey(y))
                    {
                        this.ColumnWidthsByYPos.Add(y, w);
                    }
                    else
                    {
                        if (w < this.ColumnWidthsByYPos[y]) this.ColumnWidthsByYPos[y] = w;
                    }
                    Cells.Add(rowlist[y]);
                }
            }
            int counter = 0;
            foreach (int y in this.ColumnWidthsByYPos.Keys)
            {
                ColumnIndices.Add(y, counter++);
            }            

            //the table contains all cells at their correct position
            var table = new ShapeContainer[this.RowHeights.Count, this.ColumnWidthsByYPos.Count];
            int c;
            int r = 0;
            foreach (var rowlst in tablelist.Values)
            {
                foreach (int y in rowlst.Keys)
                {
                    c = ColumnIndices[y];
                    table[r, c] = rowlst[y];
                }
                r++;
            }

            this._writer.WriteStartElement("p", "graphicFrame", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "nvGraphicFramePr", OpenXmlNamespaces.PresentationML);

            WriteCNvPr(sh.spid, "");
            //WriteCNvPr(--groupcounter, "");

            this._writer.WriteStartElement("p", "cNvGraphicFramePr", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("a", "graphicFrameLocks", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("noGrp", "1");
            this._writer.WriteEndElement(); //graphicFrameLocks
            this._writer.WriteEndElement(); //cNvGraphicFramePr

            this._writer.WriteStartElement("p", "nvPr", OpenXmlNamespaces.PresentationML);

            if (Tools.Utils.BitmaskToBool(TABLEFLAGS, 0x1 << 1))
            {
                OEPlaceHolderAtom placeholder = null;
                int exObjIdRef = -1;
                CheckClientData(group.FirstChildWithType<ShapeContainer>().FirstChildWithType<ClientData>(), ref placeholder, ref exObjIdRef);

                if (placeholder != null)
                {

                    this._writer.WriteStartElement("p", "ph", OpenXmlNamespaces.PresentationML);

                    if (!placeholder.IsObjectPlaceholder())
                    {
                        string typeValue = Utils.PlaceholderIdToXMLValue(placeholder.PlacementId);
                        this._writer.WriteAttributeString("type", typeValue);
                    }

                    switch (placeholder.PlaceholderSize)
                    {
                        case 1:
                            this._writer.WriteAttributeString("sz", "half");
                            break;
                        case 2:
                            this._writer.WriteAttributeString("sz", "quarter");
                            break;
                    }


                    if (placeholder.Position != -1)
                    {
                        this._writer.WriteAttributeString("idx", placeholder.Position.ToString());
                    }
                    else
                    {
                        try
                        {
                            var master = this._ctx.Ppt.FindMasterRecordById(group.FirstAncestorWithType<Slide>().FirstChildWithType<SlideAtom>().MasterId);
                            foreach (var cont in master.FirstChildWithType<PPDrawing>().FirstChildWithType<DrawingContainer>().FirstChildWithType<GroupContainer>().AllChildrenWithType<ShapeContainer>())
                            {
                                var s = cont.FirstChildWithType<Shape>();
                                var d = cont.FirstChildWithType<ClientData>();
                                if (d != null)
                                {
                                    var ms = new System.IO.MemoryStream(d.bytes);
                                    var rec = Record.ReadRecord(ms);
                                    if (rec is OEPlaceHolderAtom)
                                    {
                                        var placeholder2 = (OEPlaceHolderAtom)rec;
                                        if (placeholder2.PlacementId == PlaceholderEnum.MasterBody && (placeholder.PlacementId == PlaceholderEnum.Body || placeholder.PlacementId == PlaceholderEnum.Object))
                                        {
                                            if (placeholder2.Position != -1)
                                            {
                                                this._writer.WriteAttributeString("idx", placeholder2.Position.ToString());
                                            }
                                        }
                                    }
                                }
                            }

                        }

                        catch (Exception)
                        {
                            //ignore
                        }
                    }

                    this._writer.WriteEndElement();
                }
            }

            this._writer.WriteEndElement(); //nvPr
            this._writer.WriteEndElement(); //nvGraphicFramePr            

            if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
            {
                this._writer.WriteStartElement("p", "xfrm", OpenXmlNamespaces.PresentationML);

                this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                this._writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                this._writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteEndElement();
            }

            this._writer.WriteStartElement("a", "graphic", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "graphicData", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/table");
            this._writer.WriteStartElement("a", "tbl", OpenXmlNamespaces.DrawingML);
            this._writer.WriteElementString("a", "tblPr", OpenXmlNamespaces.DrawingML, "");
            this._writer.WriteStartElement("a", "tblGrid", OpenXmlNamespaces.DrawingML);

            foreach (int y in this.ColumnWidthsByYPos.Keys)
            {
                this._writer.WriteStartElement("a", "gridCol", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("w", Utils.MasterCoordToEMU(this.ColumnWidthsByYPos[y]).ToString());
                this._writer.WriteEndElement(); //gridCol
            }
            this._writer.WriteEndElement(); //tblGrid

            for (int row = 0; row < this.RowHeights.Count; row++)
            {
                this._writer.WriteStartElement("a", "tr", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("h", Utils.MasterCoordToEMU(this.RowHeights[row]).ToString());

                for (int col = 0; col < this.ColumnWidthsByYPos.Count; col++)
                {
                    this._writer.WriteStartElement("a", "tc", OpenXmlNamespaces.DrawingML);

                    if (table[row, col] != null)
                    {
                        var container = table[row, col];

                        var anch = container.FirstChildWithType<ChildAnchor>();
                        int colWidth = this.ColumnWidthsByYPos[anch.Left];

                        if (anch.rcgBounds.Height > this.RowHeights[row] && GetRowSpanCount(anch, row) > 1)
                        {
                            if (table[row + 1, col] == null)
                            {
                                this._writer.WriteAttributeString("rowSpan", GetRowSpanCount(anch, row).ToString());
                            }
                        }
                        
                        if (anch.rcgBounds.Width > colWidth)
                        {
                            this._writer.WriteAttributeString("gridSpan", GetGridSpanCount(anch, col).ToString());
                        }

                        foreach (var record in container.Children)
                        {
                            if (record is ClientTextbox)
                            {
                                this.so = container.FirstChildWithType<ShapeOptions>();
                                Apply((ClientTextbox)record, true);
                            }
                        }

                        this._writer.WriteStartElement("a", "tcPr", OpenXmlNamespaces.DrawingML);

                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dxTextLeft))
                        {
                            this._writer.WriteAttributeString("marL", this.so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op.ToString());
                        }

                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dyTextTop))
                        {
                            this._writer.WriteAttributeString("marT", this.so.OptionsByID[ShapeOptions.PropertyId.dyTextTop].op.ToString());
                        }

                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dxTextRight))
                        {
                            this._writer.WriteAttributeString("marR", this.so.OptionsByID[ShapeOptions.PropertyId.dxTextRight].op.ToString());
                        }

                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dyTextBottom))
                        {
                            this._writer.WriteAttributeString("marB", this.so.OptionsByID[ShapeOptions.PropertyId.dyTextBottom].op.ToString());
                        }

                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.anchorText))
                        {
                            switch (this.so.OptionsByID[ShapeOptions.PropertyId.anchorText].op)
                            {
                                case 0: //Top
                                    this._writer.WriteAttributeString("anchor", "t");
                                    break;
                                case 1: //Middle
                                    this._writer.WriteAttributeString("anchor", "ctr");
                                    break;
                                case 2: //Bottom
                                    this._writer.WriteAttributeString("anchor", "b");
                                    break;
                                case 3: //TopCentered
                                    this._writer.WriteAttributeString("anchor", "t");
                                    this._writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                                case 4: //MiddleCentered
                                    this._writer.WriteAttributeString("anchor", "ctr");
                                    this._writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                                case 5: //BottomCentered
                                    this._writer.WriteAttributeString("anchor", "b");
                                    this._writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                            }
                        }

                        this._writer.WriteAttributeString("horzOverflow", "overflow");

                        WriteTableLineProperties(row, col, table);

                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.FillStyleBooleanProperties))
                        {
                            var p = new FillStyleBooleanProperties(this.so.OptionsByID[ShapeOptions.PropertyId.FillStyleBooleanProperties].op);
                            if (p.fUsefFilled & p.fFilled) //  so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType))
                            {
                                new FillMapping(this._ctx, this._writer, this.parentSlideMapping).Apply(this.so);
                            }
                        }


                        this._writer.WriteEndElement(); //tcPr

                    }
                    else
                    {
                        if (col > 0 && table[row, col - 1] != null)
                        {
                            var previouscontainer = table[row, col-1];
                            var anch = previouscontainer.FirstChildWithType<ChildAnchor>();

                            if (anch.rcgBounds.Height > this.RowHeights[row] && GetRowSpanCount(anch, row) > 1)
                            {
                                this._writer.WriteAttributeString("rowSpan", GetRowSpanCount(anch, row).ToString());
                            }

                            //if (anch.rcgBounds.Height > RowHeights[row])
                            //{
                            //    _writer.WriteAttributeString("rowSpan", "2");
                            //}

                            if (anch.rcgBounds.Width > this.ColumnWidthsByYPos[anch.Left])
                            {
                                this._writer.WriteAttributeString("hMerge", "1");
                            }
                        }
                        else if (row > 0 && col > 0 && table[row - 1, col - 1] != null) //this checks the cell above on the left
                        {
                            var previouscontainer = table[row - 1, col - 1];
                            var anch = previouscontainer.FirstChildWithType<ChildAnchor>();

                            if (anch.rcgBounds.Width > this.ColumnWidthsByYPos[anch.Left])
                            {
                                this._writer.WriteAttributeString("hMerge", "1");
                            }
                        }

                        if (row > 0 && table[row-1, col] != null) //this checks the cell on the left
                        {
                            var previouscontainer = table[row-1, col];
                            var anch = previouscontainer.FirstChildWithType<ChildAnchor>();
                            int colWidth = this.ColumnWidthsByYPos[anch.Left];

                            if (anch.rcgBounds.Width > colWidth)
                            {
                                this._writer.WriteAttributeString("gridSpan", "2");
                            }

                            if (anch.rcgBounds.Height > this.RowHeights[row-1])
                            {
                                this._writer.WriteAttributeString("vMerge", "1");
                            }
                           
                        }
                        else if (row > 0 && col > 0  && table[row - 1, col - 1] != null) //this checks the cell above on the left
                        {
                            var previouscontainer = table[row - 1, col-1];
                            var anch = previouscontainer.FirstChildWithType<ChildAnchor>();

                            if (anch.rcgBounds.Height > this.RowHeights[row - 1])
                            {
                                this._writer.WriteAttributeString("vMerge", "1");
                            }
                        }

                        //insert dummy tc content
                        this._writer.WriteStartElement("a","txBody",OpenXmlNamespaces.DrawingML);
                        this._writer.WriteElementString("a","bodyPr",OpenXmlNamespaces.DrawingML,"");
                        this._writer.WriteElementString("a","lstStyle",OpenXmlNamespaces.DrawingML,"");
                        this._writer.WriteElementString("a","p",OpenXmlNamespaces.DrawingML,"");
                        this._writer.WriteEndElement(); //txBody

                        this._writer.WriteElementString("a", "tcPr", OpenXmlNamespaces.DrawingML,"");
                    }

                    this._writer.WriteEndElement(); //tc

                    //cellPointer++;
                }


                this._writer.WriteEndElement(); //tr
            }

            this._writer.WriteEndElement(); //tbl
            this._writer.WriteEndElement(); //graphicData
            this._writer.WriteEndElement(); //graphic
            this._writer.WriteEndElement(); //graphicFrame

        }

        private void WriteTableLineProperties(int row, int col, ShapeContainer[,] table)
        {
            var container = table[row, col];
            var anch = container.FirstChildWithType<ChildAnchor>();
            var so = container.FirstChildWithType<ShapeOptions>();

            ShapeContainer leftLine = null;
            ShapeContainer rightLine = null;
            ShapeContainer topLine = null;
            ShapeContainer bottomLine = null;

            foreach (int linetop in this.horizontallinelist.Keys)
            {
                var lst = this.horizontallinelist[linetop];

                if (linetop == anch.Top)
                    foreach (int lineleft in lst.Keys)
                    {
                        var line = lst[lineleft];
                        int w = line.FirstChildWithType<ChildAnchor>().rcgBounds.Width;

                        if (lineleft <= anch.Left && lineleft + w >= anch.Right) topLine = line;
                    }

                if (linetop == anch.Bottom)
                    foreach (int lineleft in lst.Keys)
                    {
                        var line = lst[lineleft];
                        int w = line.FirstChildWithType<ChildAnchor>().rcgBounds.Width;

                        if (lineleft <= anch.Left && lineleft + w >= anch.Right) bottomLine = line;
                    }
            }

            foreach (int linetop in this.verticallinelist.Keys)
            {
                var lst = this.verticallinelist[linetop];

                foreach (int lineleft in lst.Keys)
                {
                    var line = lst[lineleft];
                    int h = line.FirstChildWithType<ChildAnchor>().rcgBounds.Height;


                    if (lineleft == anch.Left)
                    //if (linetop == anch.Top) leftLine = line;
                    if (linetop <= anch.Top && linetop + h >= anch.Bottom) leftLine = line;
                    
                    if (lineleft == anch.Right)
                    //if (linetop == anch.Top) rightLine = line;
                    if (linetop <= anch.Top && linetop + h >= anch.Bottom) rightLine = line;
                }

            }



            int rows = table.GetLength(0);
            int columns = table.GetLength(1);

            ////check cell position
            //bool isLeft = (col == 0);
            //bool isRight = (col == columns - 1);
            //bool isTop = (row == 0);
            //bool isBottom = (row == rows - 1);

            int span;
            int colWidth = this.ColumnWidthsByYPos[anch.Left];
            if (anch.rcgBounds.Height > this.RowHeights[row])
            {
                //recheck isBottom
                span = GetRowSpanCount(anch, row);
                //isBottom = (row + span - 1 == rows - 1);
            }
            if (anch.rcgBounds.Width > colWidth)
            {
                //recheck isRight
                span = GetGridSpanCount(anch, col);
                //isRight = (col + span - 1 == columns - 1);
            }

            this._writer.WriteStartElement("a", "lnL", OpenXmlNamespaces.DrawingML);
            WriteLineProperties(leftLine, so);
            //if (isLeft) WriteLineProperties(outerLine, so); else WriteLineProperties(innerLine, null); 
            this._writer.WriteEndElement(); //lnL

            this._writer.WriteStartElement("a", "lnR", OpenXmlNamespaces.DrawingML);
            WriteLineProperties(rightLine, so);
            //if (isRight) WriteLineProperties(outerLine, so); else WriteLineProperties(innerLine, null);
            this._writer.WriteEndElement(); //lnR

            this._writer.WriteStartElement("a", "lnT", OpenXmlNamespaces.DrawingML);
            WriteLineProperties(topLine, so);
            //if (isTop) WriteLineProperties(outerLine, so); else WriteLineProperties(innerLine, null);
            this._writer.WriteEndElement(); //lnT

            this._writer.WriteStartElement("a", "lnB", OpenXmlNamespaces.DrawingML);
            WriteLineProperties(bottomLine, so);
            //if (isBottom) WriteLineProperties(outerLine, so); else WriteLineProperties(innerLine, null);
            this._writer.WriteEndElement(); //lnB

            this._writer.WriteStartElement("a", "lnTlToBr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteElementString("a", "noFill", OpenXmlNamespaces.DrawingML, "");
            this._writer.WriteEndElement(); //lnTlToBr

            this._writer.WriteStartElement("a", "lnBlToTr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteElementString("a", "noFill", OpenXmlNamespaces.DrawingML, "");
            this._writer.WriteEndElement(); //lnBlToTr
                      
        }

        private void WriteLineProperties(ShapeOptions soline)
        {
            var fsb = new LineStyleBooleans(soline.OptionsByID[ShapeOptions.PropertyId.lineStyleBooleans].op);
            if (fsb.fUsefLine && fsb.fLine)
            {
                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor))
                {
                    this._writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    string SchemeType = "";
                    string colorVal = Utils.getRGBColorFromOfficeArtCOLORREF(soline.OptionsByID[ShapeOptions.PropertyId.lineColor].op, soline.FirstAncestorWithType<Slide>(), this.so, ref SchemeType);

                    if (SchemeType.Length == 0)
                    {
                        this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", colorVal);
                    }
                    else
                    {
                        this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", SchemeType);
                    }
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                }
                else
                {
                    this._writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", "000000");
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                }
            }
        }

        private void WriteLineProperties(ShapeContainer lineCont, ShapeOptions soframe)
        {
            if (lineCont == null)
            {
                this._writer.WriteElementString("a", "noFill", OpenXmlNamespaces.DrawingML, "");
                return;
            }

            foreach (var soline in lineCont.AllChildrenWithType<ShapeOptions>())
            {
                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineWidth))
                {
                    uint w = soline.OptionsByID[ShapeOptions.PropertyId.lineWidth].op;
                    this._writer.WriteAttributeString("w", w.ToString());
                }
                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndCapStyle))
                {
                    //switch (soline.OptionsByID[ShapeOptions.PropertyId.lineEndCapStyle].op)
                    //{
                    //    case 0: //round
                    //        _writer.WriteAttributeString("cap", "rnd");
                    //        break;
                    //    case 1: //square
                    //        _writer.WriteAttributeString("cap", "sq");
                    //        break;
                    //    case 2: //flat
                    this._writer.WriteAttributeString("cap", "flat");
                            //break;
                    //}
                }

                this._writer.WriteAttributeString("cmpd", "sng");
                this._writer.WriteAttributeString("algn", "ctr");


                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStyleBooleans))
                {
                    WriteLineProperties(soline);
                    //LineStyleBooleans fsb = new LineStyleBooleans(soline.OptionsByID[ShapeOptions.PropertyId.lineStyleBooleans].op);
                    //if (fsb.fUsefLine && fsb.fLine)
                    //{
                    //    if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor))
                    //    {
                    //        _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    //        string SchemeType = "";
                    //        string colorVal = Utils.getRGBColorFromOfficeArtCOLORREF(soline.OptionsByID[ShapeOptions.PropertyId.lineColor].op, lineCont.FirstAncestorWithType<Slide>(), so, ref SchemeType);

                    //        if (SchemeType.Length == 0)
                    //        {
                    //            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    //            _writer.WriteAttributeString("val", colorVal);
                    //        }
                    //        else
                    //        {
                    //            _writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                    //            _writer.WriteAttributeString("val", SchemeType);
                    //        }
                    //        _writer.WriteEndElement();
                    //        _writer.WriteEndElement();
                    //    }
                    //    else
                    //    {
                    //        _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    //        _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    //        _writer.WriteAttributeString("val", "000000");
                    //        _writer.WriteEndElement();
                    //        _writer.WriteEndElement();
                    //    }
                    //}
                } else if (soframe != null && soframe.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStyleBooleans))
                {
                    WriteLineProperties(soframe);
                    //LineStyleBooleans fsb = new LineStyleBooleans(soframe.OptionsByID[ShapeOptions.PropertyId.lineStyleBooleans].op);
                    //if (fsb.fUsefLine && fsb.fLine)
                    //{
                    //     if (soframe.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor))
                    //    {
                    //        _writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    //        string SchemeType = "";
                    //        string colorVal = Utils.getRGBColorFromOfficeArtCOLORREF(soframe.OptionsByID[ShapeOptions.PropertyId.lineColor].op, lineCont.FirstAncestorWithType<Slide>(), soframe, ref SchemeType);

                    //        if (SchemeType.Length == 0)
                    //        {
                    //            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    //            _writer.WriteAttributeString("val", colorVal);
                    //        }
                    //        else
                    //        {
                    //            _writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                    //            _writer.WriteAttributeString("val", SchemeType);
                    //        }
                    //        _writer.WriteEndElement();
                    //        _writer.WriteEndElement();
                    //    }
                    //}
                }
               
                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineDashing))
                {
                    this._writer.WriteStartElement("a", "prstDash", OpenXmlNamespaces.DrawingML);
                    switch ((ShapeOptions.LineDashing)soline.OptionsByID[ShapeOptions.PropertyId.lineDashing].op)
                    {
                        case ShapeOptions.LineDashing.Solid:
                            this._writer.WriteAttributeString("val", "solid");
                            break;
                        case ShapeOptions.LineDashing.DashSys:
                            this._writer.WriteAttributeString("val", "sysDash");
                            break;
                        case ShapeOptions.LineDashing.DotSys:
                            this._writer.WriteAttributeString("val", "sysDot");
                            break;
                        case ShapeOptions.LineDashing.DashDotSys:
                            this._writer.WriteAttributeString("val", "sysDashDot");
                            break;
                        case ShapeOptions.LineDashing.DashDotDotSys:
                            this._writer.WriteAttributeString("val", "sysDashDotDot");
                            break;
                        case ShapeOptions.LineDashing.DotGEL:
                            this._writer.WriteAttributeString("val", "dot");
                            break;
                        case ShapeOptions.LineDashing.DashGEL:
                            this._writer.WriteAttributeString("val", "dash");
                            break;
                        case ShapeOptions.LineDashing.LongDashGEL:
                            this._writer.WriteAttributeString("val", "lgDash");
                            break;
                        case ShapeOptions.LineDashing.DashDotGEL:
                            this._writer.WriteAttributeString("val", "dashDot");
                            break;
                        case ShapeOptions.LineDashing.LongDashDotGEL:
                            this._writer.WriteAttributeString("val", "lgDashDot");
                            break;
                        case ShapeOptions.LineDashing.LongDashDotDotGEL:
                            this._writer.WriteAttributeString("val", "lgDashDotDot");
                            break;
                    }
                    this._writer.WriteEndElement();
                }
                else if (soframe != null && soframe.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineDashing))
                {
                    this._writer.WriteStartElement("a", "prstDash", OpenXmlNamespaces.DrawingML);
                    switch ((ShapeOptions.LineDashing)soframe.OptionsByID[ShapeOptions.PropertyId.lineDashing].op)
                    {
                        case ShapeOptions.LineDashing.Solid:
                            this._writer.WriteAttributeString("val", "solid");
                            break;
                        case ShapeOptions.LineDashing.DashSys:
                            this._writer.WriteAttributeString("val", "sysDash");
                            break;
                        case ShapeOptions.LineDashing.DotSys:
                            this._writer.WriteAttributeString("val", "sysDot");
                            break;
                        case ShapeOptions.LineDashing.DashDotSys:
                            this._writer.WriteAttributeString("val", "sysDashDot");
                            break;
                        case ShapeOptions.LineDashing.DashDotDotSys:
                            this._writer.WriteAttributeString("val", "sysDashDotDot");
                            break;
                        case ShapeOptions.LineDashing.DotGEL:
                            this._writer.WriteAttributeString("val", "dot");
                            break;
                        case ShapeOptions.LineDashing.DashGEL:
                            this._writer.WriteAttributeString("val", "dash");
                            break;
                        case ShapeOptions.LineDashing.LongDashGEL:
                            this._writer.WriteAttributeString("val", "lgDash");
                            break;
                        case ShapeOptions.LineDashing.DashDotGEL:
                            this._writer.WriteAttributeString("val", "dashDot");
                            break;
                        case ShapeOptions.LineDashing.LongDashDotGEL:
                            this._writer.WriteAttributeString("val", "lgDashDot");
                            break;
                        case ShapeOptions.LineDashing.LongDashDotDotGEL:
                            this._writer.WriteAttributeString("val", "lgDashDotDot");
                            break;
                    }
                    this._writer.WriteEndElement();
                }

                this._writer.WriteElementString("a", "round", OpenXmlNamespaces.DrawingML, "");

                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStartArrowhead))
                {
                    var val = (ShapeOptions.LineEnd)soline.OptionsByID[ShapeOptions.PropertyId.lineStartArrowhead].op;
                    if (val != ShapeOptions.LineEnd.NoEnd)
                    {
                        this._writer.WriteStartElement("a", "headEnd", OpenXmlNamespaces.DrawingML);
                        switch (val)
                        {
                            case ShapeOptions.LineEnd.ArrowEnd:
                                this._writer.WriteAttributeString("type", "triangle");
                                break;
                            case ShapeOptions.LineEnd.ArrowStealthEnd:
                                this._writer.WriteAttributeString("type", "stealth");
                                break;
                            case ShapeOptions.LineEnd.ArrowDiamondEnd:
                                this._writer.WriteAttributeString("type", "diamond");
                                break;
                            case ShapeOptions.LineEnd.ArrowOvalEnd:
                                this._writer.WriteAttributeString("type", "oval");
                                break;
                            case ShapeOptions.LineEnd.ArrowOpenEnd:
                                this._writer.WriteAttributeString("type", "arrow");
                                break;
                            case ShapeOptions.LineEnd.ArrowChevronEnd: //this should be ignored
                            case ShapeOptions.LineEnd.ArrowDoubleChevronEnd:
                                this._writer.WriteAttributeString("type", "triangle");
                                break;
                        }
                        this._writer.WriteAttributeString("w", "med");
                        this._writer.WriteAttributeString("len", "med");
                        this._writer.WriteEndElement(); //headEnd
                    }
                }
                else
                {
                    this._writer.WriteStartElement("a", "headEnd", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("type", "none");
                    this._writer.WriteAttributeString("w", "med");
                    this._writer.WriteAttributeString("len", "med");
                    this._writer.WriteEndElement(); //headEnd
                }

                if (soline.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndArrowhead))
                {
                    var val = (ShapeOptions.LineEnd)soline.OptionsByID[ShapeOptions.PropertyId.lineEndArrowhead].op;
                    if (val != ShapeOptions.LineEnd.NoEnd)
                    {
                        this._writer.WriteStartElement("a", "tailEnd", OpenXmlNamespaces.DrawingML);
                        switch (val)
                        {
                            case ShapeOptions.LineEnd.ArrowEnd:
                                this._writer.WriteAttributeString("type", "triangle");
                                break;
                            case ShapeOptions.LineEnd.ArrowStealthEnd:
                                this._writer.WriteAttributeString("type", "stealth");
                                break;
                            case ShapeOptions.LineEnd.ArrowDiamondEnd:
                                this._writer.WriteAttributeString("type", "diamond");
                                break;
                            case ShapeOptions.LineEnd.ArrowOvalEnd:
                                this._writer.WriteAttributeString("type", "oval");
                                break;
                            case ShapeOptions.LineEnd.ArrowOpenEnd:
                                this._writer.WriteAttributeString("type", "arrow");
                                break;
                            case ShapeOptions.LineEnd.ArrowChevronEnd: //this should be ignored
                            case ShapeOptions.LineEnd.ArrowDoubleChevronEnd:
                                this._writer.WriteAttributeString("type", "triangle");
                                break;
                        }
                        this._writer.WriteAttributeString("w", "med");
                        this._writer.WriteAttributeString("len", "med");
                        this._writer.WriteEndElement(); //tailnd
                    }
                }
                else
                {
                    this._writer.WriteStartElement("a", "tailEnd", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("type", "none");
                    this._writer.WriteAttributeString("w", "med");
                    this._writer.WriteAttributeString("len", "med");
                    this._writer.WriteEndElement(); //tailEnd
                }

                break;  
            }
        }


        private ShapeOptions so;
        public void Apply(ShapeContainer container)
        {
            Apply(container, "","", "");
        }
        public void Apply(ShapeContainer container, string footertext, string headertext, string datetext)
        {

            this._footertext = footertext;
            this._headertext = headertext;
            this._datetext = datetext;
            var clientData = container.FirstChildWithType<ClientData>();

            RegularContainer slide = container.FirstAncestorWithType<Slide>();
            if (slide == null) slide = container.FirstAncestorWithType<Note>();
            if (slide == null) slide = container.FirstAncestorWithType<Handout>();

            bool continueShape = true;

            var sh = container.FirstChildWithType<Shape>();

            this.so = container.FirstChildWithType<ShapeOptions>();
            //if (clientData == null)
                if (this.so != null)
                {
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.Pib))
                    {
                        if (sh.fOleShape)
                        {
                            OEPlaceHolderAtom placeholder = null;
                            int exObjIdRef = -1;
                            CheckClientData(container.FirstChildWithType<ClientData>(), ref placeholder, ref exObjIdRef);
                            var oleContainer = this._ctx.Ppt.OleObjects[exObjIdRef];
                            if (oleContainer.FirstChildWithType<CStringAtom>() != null)
                            if (oleContainer.FirstChildWithType<CStringAtom>().Text == "Chart")
                            {
                                writeOle(container, oleContainer);
                                continueShape = false;
                            }
                        }

                        if (continueShape)
                        {
                            writePic(container);
                            continueShape = false;
                        }
                    }
                }
                else
                {
                this.so =  new ShapeOptions();
                }

            ShapeOptions sndSo = null;
            string prstGeom = "";
            if (container.AllChildrenWithType<ShapeOptions>().Count > 1)
            {
                sndSo = ((RegularContainer)sh.ParentRecord).AllChildrenWithType<ShapeOptions>()[1];
                if (sndSo.OptionsByID.ContainsKey(ShapeOptions.PropertyId.metroBlob))
                {
                    IZipReader reader = null;
                    try
                    {
                        var metroBlob = sndSo.OptionsByID[ShapeOptions.PropertyId.metroBlob];
                        var code = metroBlob.opComplex;
                        string path = Path.GetTempFileName();
                        var fs = new FileStream(path, FileMode.Create);
                        fs.Write(code, 0, code.Length);
                        fs.Close();

                        reader = ZipFactory.OpenArchive(path);
                        var mems = new StreamReader(reader.GetEntry("drs/shapexml.xml"));
                        string xml = mems.ReadToEnd();
                        xml = Tools.Utils.replaceOutdatedNamespaces(xml);
                        //xml = xml.Substring(xml.IndexOf("<p:sp")); //remove xml declaration

                        xml = xml.Substring(xml.IndexOf("<a:prstGeom"));
                        if (xml.Length > 0)
                        {
                            xml = xml.Substring(0, xml.IndexOf("</a:prstGeom>") + 13);
                            prstGeom = xml;
                        }

                        //_writer.WriteRaw(xml);
                        //continueShape = false;
                        reader.Close();
                    }
                    catch (Exception)
                    {
                        continueShape = true;
                        if (reader != null) reader.Close();
                    }
                }
            }

            if (continueShape)
            {
                if (sh.fConnector)
                {
                    string idStart = "";
                    string idEnd = "";
                    string idxStart = "0";
                    string idxEnd = "0";
                    foreach (var rule in container.FirstAncestorWithType<DrawingContainer>().FirstChildWithType<SolverContainer>().AllChildrenWithType<FConnectorRule>())
                    {
                        if (rule.spidC == sh.spid) //spidC marks the connector shape
                        {
                            if (!this.spidToId.ContainsKey((int)rule.spidA))
                            {
                                this.spidToId.Add((int)rule.spidA, ++this._idCnt);
                            }
                            if (!this.spidToId.ContainsKey((int)rule.spidB))
                            {
                                this.spidToId.Add((int)rule.spidB, ++this._idCnt);
                            }

                            idStart = this.spidToId[(int)rule.spidA].ToString(); //spidA marks the start shape
                            idEnd = this.spidToId[(int)rule.spidB].ToString(); //spidB marks the end shape
                            idxStart = rule.cptiA.ToString();
                            idxEnd = rule.cptiB.ToString();
                        }
                    }

                    this._writer.WriteStartElement("p", "cxnSp", OpenXmlNamespaces.PresentationML);
                    this._writer.WriteStartElement("p", "nvCxnSpPr", OpenXmlNamespaces.PresentationML);
                    WriteCNvPr(sh.spid, "");

                    this._writer.WriteStartElement("p", "cNvCxnSpPr", OpenXmlNamespaces.PresentationML);

                    this._writer.WriteStartElement("a", "stCxn", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("id", idStart);
                    this._writer.WriteAttributeString("idx", idxStart);
                    this._writer.WriteEndElement(); //stCxn

                    this._writer.WriteStartElement("a", "endCxn", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("id", idEnd);
                    this._writer.WriteAttributeString("idx", idxEnd);
                    this._writer.WriteEndElement(); //endCxn

                    this._writer.WriteEndElement(); //cNvCxnSpPr
                }
                else
                {
                    this._writer.WriteStartElement("p", "sp", OpenXmlNamespaces.PresentationML);
                    this._writer.WriteStartElement("p", "nvSpPr", OpenXmlNamespaces.PresentationML);
                    WriteCNvPr(sh.spid, "");

                    this._writer.WriteElementString("p", "cNvSpPr", OpenXmlNamespaces.PresentationML, "");
                }

                this._writer.WriteStartElement("p", "nvPr", OpenXmlNamespaces.PresentationML);

                OEPlaceHolderAtom placeholder = null;
                int exObjIdRef = 0;
                CheckClientData(clientData, ref placeholder, ref exObjIdRef);

                this._writer.WriteEndElement();

                this._writer.WriteEndElement();


                // Visible shape properties
                this._writer.WriteStartElement("p", "spPr", OpenXmlNamespaces.PresentationML);

                var anchor = container.FirstChildWithType<ClientAnchor>();
                var chAnchor = container.FirstChildWithType<ChildAnchor>();

                bool swapHeightWidth = false;
                Double dc = 0;
                if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
                {
                    this._writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);
                    if (sh.fFlipH) this._writer.WriteAttributeString("flipH", "1");
                    if (sh.fFlipV) this._writer.WriteAttributeString("flipV", "1");
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.rotation))
                    {
                        var bytes = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.rotation].op);
                        int integral = BitConverter.ToInt16(bytes, 2);
                        uint fractional = BitConverter.ToUInt16(bytes, 0);
                        decimal result = integral +((decimal)fractional / (decimal)65536);

                        Double w = anchor.Bottom - anchor.Top;
                        Double h = anchor.Right - anchor.Left;

                        dc = (w - h) / 2;

                        if (Math.Abs(result) > 45 && Math.Abs(result) < 135) swapHeightWidth = true;
                        if (Math.Abs(result) > 225 && Math.Abs(result) < 315) swapHeightWidth = true;

                        //if (result < 0 && sh.fFlipH == false) result = result * -1;

                        string rotation = Math.Floor(result * 60000).ToString();
                        if (result != 0)
                        {
                            this._writer.WriteAttributeString("rot", rotation);
                        }

                    }

                    if (container.FirstAncestorWithType<GroupContainer>().FirstAncestorWithType<GroupContainer>() == null)
                    {
                        this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);

                        if (swapHeightWidth)
                        {
                            this._writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left - (int)dc).ToString());
                            this._writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top + (int)dc).ToString());
                        }
                        else
                        {
                            this._writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                            this._writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                        }
                        this._writer.WriteEndElement();

                        this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);

                        if (swapHeightWidth)
                        {
                            this._writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                            this._writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());                            
                        }
                        else
                        {
                            this._writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                            this._writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                        }
                        this._writer.WriteEndElement();
                    }
                    else
                    {
                        var clanchor = container.FirstChildWithType<ClientAnchor>();
                        if (clanchor == null)
                        {
                            this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                            this._writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                            this._writer.WriteEndElement();

                            this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                            if (swapHeightWidth)
                            {
                                this._writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                                this._writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                            }
                            else
                            {
                                this._writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                                this._writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                            }
                            this._writer.WriteEndElement();
                        }
                        else
                        {
                            this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("x", (anchor.Left).ToString());
                            this._writer.WriteAttributeString("y", (anchor.Top).ToString());
                            this._writer.WriteEndElement();

                            this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                            if (swapHeightWidth)
                            {
                                this._writer.WriteAttributeString("cx", (anchor.Bottom - anchor.Top).ToString());
                                this._writer.WriteAttributeString("cy", (anchor.Right - anchor.Left).ToString());
                            }
                            else
                            {
                                this._writer.WriteAttributeString("cx", (anchor.Right - anchor.Left).ToString());
                                this._writer.WriteAttributeString("cy", (anchor.Bottom - anchor.Top).ToString());
                            }
                            this._writer.WriteEndElement();
                        }
                    }
                    this._writer.WriteEndElement();
                }
                else if (chAnchor != null && chAnchor.Right >= chAnchor.Left && chAnchor.Bottom >= chAnchor.Top)
                {
                    var groupAnchor = container.FirstAncestorWithType<GroupContainer>().FirstChildWithType<ShapeContainer>().FirstChildWithType<ClientAnchor>();
                    var rec = container.FirstAncestorWithType<GroupContainer>().FirstChildWithType<ShapeContainer>().FirstChildWithType<GroupShapeRecord>().rcgBounds;

                    this._writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.rotation))
                    {
                        var bytes = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.rotation].op);
                        int integral = BitConverter.ToInt16(bytes, 2);
                        uint fractional = BitConverter.ToUInt16(bytes, 0);
                        decimal result = integral + ((decimal)fractional / (decimal)65536);

                        //if (result < 0 && sh.fFlipH == false) result = result * -1;

                        Double w = chAnchor.Bottom - chAnchor.Top;
                        Double h = chAnchor.Right - chAnchor.Left;

                        dc = (w - h) / 2;

                        if (Math.Abs(result) > 45 && Math.Abs(result) < 135) swapHeightWidth = true;
                        if (Math.Abs(result) > 225 && Math.Abs(result) < 315) swapHeightWidth = true;

                        string rotation = Math.Floor(result * 60000).ToString();
                        this._writer.WriteAttributeString("rot", rotation);
                    }

                    if (sh.fFlipH) this._writer.WriteAttributeString("flipH", "1");
                    if (sh.fFlipV) this._writer.WriteAttributeString("flipV", "1");

                    this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                    if (swapHeightWidth)
                    {
                        this._writer.WriteAttributeString("x", (chAnchor.Left - (int)dc).ToString());
                        this._writer.WriteAttributeString("y", (chAnchor.Top + (int)dc).ToString());
                    }
                    else
                    {
                        this._writer.WriteAttributeString("x", (chAnchor.Left).ToString());
                        this._writer.WriteAttributeString("y", (chAnchor.Top).ToString());
                    }
                    this._writer.WriteEndElement();

                    this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                    if (swapHeightWidth)
                    {
                        this._writer.WriteAttributeString("cx", (chAnchor.Bottom - chAnchor.Top).ToString());
                        this._writer.WriteAttributeString("cy", (chAnchor.Right - chAnchor.Left).ToString());
                    }
                    else
                    {
                        this._writer.WriteAttributeString("cx", (chAnchor.Right - chAnchor.Left).ToString());
                        this._writer.WriteAttributeString("cy", (chAnchor.Bottom - chAnchor.Top).ToString());
                    }
                    this._writer.WriteEndElement();

                    this._writer.WriteEndElement();
                }

                if (prstGeom.Length > 0) //this means a predefined shape in a blob
                {
                    this._writer.WriteRaw(prstGeom);
                }
                else if (sh.Instance != 0 & !this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pSegmentInfo)) //this means a predefined shape
                {
                    WriteprstGeom(sh);
                }
                else //this means a custom shape
                {
                    WritecustGeom(sh);
                }

                WriteShapeProperties(sh, placeholder != null, slide);

                this._writer.WriteEndElement();

                bool TextBoxFound = false;

                // Descend into unsupported records
                foreach (var record in container.Children)
                {
                    DynamicApply(record);
                    if (record is ClientTextbox) TextBoxFound = true;
                }

                
                if (!TextBoxFound & !sh.fConnector)
                {

                    this._writer.WriteStartElement("p", "txBody", OpenXmlNamespaces.PresentationML);
                    writeBodyPr(container, false, false);
                    this._writer.WriteElementString("a", "lstStyle", OpenXmlNamespaces.DrawingML, "");
                    this._writer.WriteStartElement("a", "p", OpenXmlNamespaces.DrawingML);

                    this._writer.WriteStartElement("a", "pPr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("algn", "ctr");
                    this._writer.WriteAttributeString("fontAlgn", "auto");
                    this._writer.WriteEndElement();

                    //check if there is text in so
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.gtextUNICODE))
                    {
                        var bytes = this.so.OptionsByID[ShapeOptions.PropertyId.gtextUNICODE].opComplex;
                        string sText = Encoding.Unicode.GetString(bytes);
                        if (sText.Contains("\0")) sText = sText.Substring(0, sText.IndexOf("\0"));
                        this._writer.WriteStartElement("a", "r", OpenXmlNamespaces.DrawingML);

                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.gtextFont))
                        {
                            bytes = this.so.OptionsByID[ShapeOptions.PropertyId.gtextFont].opComplex;
                            string sFont = Encoding.Unicode.GetString(bytes);
                            if (sFont.Contains("\0")) sFont = sFont.Substring(0, sFont.IndexOf("\0"));

                            this._writer.WriteStartElement("a", "rPr", OpenXmlNamespaces.DrawingML);
                            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.gtextSize))
                            {
                                this._writer.WriteAttributeString("sz", (this.so.OptionsByID[ShapeOptions.PropertyId.gtextSize].op / 0x100).ToString());
                            }
                            else
                            {
                                this._writer.WriteAttributeString("sz", "3600");
                            }
                            
                            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.GeometryTextBooleanProperties))
                            {
                                var gb = new GeometryTextBooleanProperties(this.so.OptionsByID[ShapeOptions.PropertyId.GeometryTextBooleanProperties].op);
                                if (gb.fUsegtextFKern & gb.gtextFKern)
                                {
                                    this._writer.WriteAttributeString("kern", "10");
                                }

                                if (gb.fUsegtextFItalic & gb.gtextFItalic)
                                {
                                    this._writer.WriteAttributeString("i", "1");
                                }
                            }

                            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStyleBooleans))
                            {
                                var lb = new LineStyleBooleans(this.so.OptionsByID[ShapeOptions.PropertyId.lineStyleBooleans].op);
                                if (lb.fUsefLine & lb.fLine)
                                {
                                    this._writer.WriteStartElement("a", "ln", OpenXmlNamespaces.DrawingML);

                                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineWidth))
                                    {
                                        uint w = this.so.OptionsByID[ShapeOptions.PropertyId.lineWidth].op;
                                        this._writer.WriteAttributeString("w", w.ToString());
                                    }

                                    string color = "000000";
                                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor))
                                    {
                                        color = Utils.getRGBColorFromOfficeArtCOLORREF(this.so.OptionsByID[ShapeOptions.PropertyId.lineColor].op, slide, this.so);
                                    }

                                    bool ignoreColor = false;
                                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.ThreeDObjectBooleanProperties))
                                    {
                                        var tdo = new ThreeDObjectProperties(this.so.OptionsByID[ShapeOptions.PropertyId.ThreeDObjectBooleanProperties].op);

                                        if (tdo.fc3D && tdo.fUsefc3D)
                                        {
                                            ignoreColor = true;
                                        }
                                    }

                                    if (!ignoreColor)
                                    {
                                        this._writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                                        this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                                        this._writer.WriteAttributeString("val", color);
                                        this._writer.WriteEndElement();
                                        this._writer.WriteEndElement();
                                    }
                                    this._writer.WriteEndElement();
                                }
                            }

                            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.FillStyleBooleanProperties))
                            {
                                var p = new FillStyleBooleanProperties(this.so.OptionsByID[ShapeOptions.PropertyId.FillStyleBooleanProperties].op);
                                if (p.fUsefFilled & p.fFilled)
                                {
                                    new FillMapping(this._ctx, this._writer, this.parentSlideMapping).Apply(this.so);
                                }
                            }

                            //shadow
                            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.ShadowStyleBooleanProperties))
                            {
                                var sp = new ShadowStyleBooleanProperties(this.so.OptionsByID[ShapeOptions.PropertyId.ShadowStyleBooleanProperties].op);
                                if (sp.fUsefShadow & sp.fShadow)
                                {
                                    new ShadowMapping(this._ctx, this._writer).Apply(this.so);
                                }
                            }

                            this._writer.WriteStartElement("a", "latin", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("typeface", sFont);
                            this._writer.WriteEndElement();
                            this._writer.WriteEndElement();
                        }

                        this._writer.WriteElementString("a", "t", OpenXmlNamespaces.DrawingML, sText);
                        this._writer.WriteEndElement();
                    }

                    this._writer.WriteElementString("a", "endParaRPr", OpenXmlNamespaces.DrawingML, "");
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                }

                this._writer.WriteEndElement();
            }
        }

        private void WriteShapeProperties(Shape sh, bool isPlaceholder, RegularContainer slide)
        {
            var container = (ShapeContainer)sh.ParentRecord;

            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.FillStyleBooleanProperties))
            {
                var p = new FillStyleBooleanProperties(this.so.OptionsByID[ShapeOptions.PropertyId.FillStyleBooleanProperties].op);
                if (p.fUsefFilled & p.fFilled) //  so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType))
                {
                    new FillMapping(this._ctx, this._writer, this.parentSlideMapping).Apply(this.so);
                }
            }
            else if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
            {
                if (sh.Instance != 0xca & isPlaceholder == false)
                {
                    string colorval = Utils.getRGBColorFromOfficeArtCOLORREF(this.so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, this.so);
                    this._writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", colorval);
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity) && this.so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op != 65536)
                    {
                        this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", Math.Round(((decimal)this.so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                        this._writer.WriteEndElement();
                    }
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                }
            }

            this._writer.WriteStartElement("a", "ln", OpenXmlNamespaces.DrawingML);
            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineWidth))
            {
                this._writer.WriteAttributeString("w", this.so.OptionsByID[ShapeOptions.PropertyId.lineWidth].op.ToString());
            }
            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndCapStyle))
            {
                switch (this.so.OptionsByID[ShapeOptions.PropertyId.lineEndCapStyle].op)
                {
                    case 0: //round
                        this._writer.WriteAttributeString("cap", "rnd");
                        break;
                    case 1: //square
                        this._writer.WriteAttributeString("cap", "sq");
                        break;
                    case 2: //flat
                        this._writer.WriteAttributeString("cap", "flat");
                        break;
                }
            }


            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineType))
            {
                switch (this.so.OptionsByID[ShapeOptions.PropertyId.lineType].op)
                {
                    case 0: //solid
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStyleBooleans))
                        {
                            var lineStyle = new LineStyleBooleans(this.so.OptionsByID[ShapeOptions.PropertyId.lineStyleBooleans].op);
                            if (lineStyle.fLine)
                            {
                                string colorval = "000000";
                                if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor))
                                    colorval = Utils.getRGBColorFromOfficeArtCOLORREF(this.so.OptionsByID[ShapeOptions.PropertyId.lineColor].op, slide, this.so);
                                this._writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("val", colorval);
                                this._writer.WriteEndElement();
                                this._writer.WriteEndElement();
                            }
                        }
                        break;
                    case 1: //pattern
                        uint blipIndex = this.so.OptionsByID[ShapeOptions.PropertyId.lineFillBlip].op;
                        var gr = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
                        var bse = (BlipStoreEntry)gr.FirstChildWithType<BlipStoreContainer>().Children[(int)blipIndex - 1];
                        var b = (BitmapBlip)this._ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

                        this._writer.WriteStartElement("a", "pattFill", OpenXmlNamespaces.DrawingML);

                        this._writer.WriteAttributeString("prst", Utils.getPrstForPatternCode(b.m_bTag)); //Utils.getPrstForPattern(blipNamePattern));

                        this._writer.WriteStartElement("a", "fgClr", OpenXmlNamespaces.DrawingML);

                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor))
                        {
                            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", Utils.getRGBColorFromOfficeArtCOLORREF(this.so.OptionsByID[ShapeOptions.PropertyId.lineColor].op, container.FirstAncestorWithType<Slide>(), this.so));
                            this._writer.WriteEndElement();
                        }
                        else
                        {
                            this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", "tx1");
                            this._writer.WriteEndElement();
                        }

                        this._writer.WriteEndElement();

                        this._writer.WriteStartElement("a", "bgClr", OpenXmlNamespaces.DrawingML);

                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineBackColor))
                        {
                            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", Utils.getRGBColorFromOfficeArtCOLORREF(this.so.OptionsByID[ShapeOptions.PropertyId.lineBackColor].op, container.FirstAncestorWithType<Slide>(), this.so));
                            this._writer.WriteEndElement();
                        }
                        else
                        {
                            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", "FFFFFF");
                            this._writer.WriteEndElement();
                        }

                        this._writer.WriteEndElement();

                        this._writer.WriteEndElement();

                        break;
                    case 2: //texture
                        break;
                    default:
                        break;
                }
            }
            else
            {
                if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStyleBooleans))
                {
                    var lineStyle = new LineStyleBooleans(this.so.OptionsByID[ShapeOptions.PropertyId.lineStyleBooleans].op);
                    if (lineStyle.fLine)
                    {
                        string colorval = "000000";
                        string schemeType = "";
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineColor))
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(this.so.OptionsByID[ShapeOptions.PropertyId.lineColor].op, slide, this.so, ref schemeType);
                        this._writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);

                        if (schemeType.Length == 0)
                        {
                            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", colorval);
                        }
                        else
                        {
                            this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", schemeType);
                        }
                                             
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineOpacity) && this.so.OptionsByID[ShapeOptions.PropertyId.lineOpacity].op != 65536)
                        {
                            this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", Math.Round(((decimal)this.so.OptionsByID[ShapeOptions.PropertyId.lineOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            this._writer.WriteEndElement();
                        }
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                    }
                }
            }

            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineDashing))
            {
                this._writer.WriteStartElement("a", "prstDash", OpenXmlNamespaces.DrawingML);
                switch ((ShapeOptions.LineDashing)this.so.OptionsByID[ShapeOptions.PropertyId.lineDashing].op)
                {
                    case ShapeOptions.LineDashing.Solid:
                        this._writer.WriteAttributeString("val", "solid");
                        break;
                    case ShapeOptions.LineDashing.DashSys:
                        this._writer.WriteAttributeString("val", "sysDash");
                        break;
                    case ShapeOptions.LineDashing.DotSys:
                        this._writer.WriteAttributeString("val", "sysDot");
                        break;
                    case ShapeOptions.LineDashing.DashDotSys:
                        this._writer.WriteAttributeString("val", "sysDashDot");
                        break;
                    case ShapeOptions.LineDashing.DashDotDotSys:
                        this._writer.WriteAttributeString("val", "sysDashDotDot");
                        break;
                    case ShapeOptions.LineDashing.DotGEL:
                        this._writer.WriteAttributeString("val", "dot");
                        break;
                    case ShapeOptions.LineDashing.DashGEL:
                        this._writer.WriteAttributeString("val", "dash");
                        break;
                    case ShapeOptions.LineDashing.LongDashGEL:
                        this._writer.WriteAttributeString("val", "lgDash");
                        break;
                    case ShapeOptions.LineDashing.DashDotGEL:
                        this._writer.WriteAttributeString("val", "dashDot");
                        break;
                    case ShapeOptions.LineDashing.LongDashDotGEL:
                        this._writer.WriteAttributeString("val", "lgDashDot");
                        break;
                    case ShapeOptions.LineDashing.LongDashDotDotGEL:
                        this._writer.WriteAttributeString("val", "lgDashDotDot");
                        break;
                }
                this._writer.WriteEndElement();
            }

            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStartArrowhead))
            {
                var val = (ShapeOptions.LineEnd)this.so.OptionsByID[ShapeOptions.PropertyId.lineStartArrowhead].op;
                if (val != ShapeOptions.LineEnd.NoEnd)
                {
                    this._writer.WriteStartElement("a", "headEnd", OpenXmlNamespaces.DrawingML);
                    switch (val)
                    {
                        case ShapeOptions.LineEnd.ArrowEnd:
                            this._writer.WriteAttributeString("type", "triangle");
                            break;
                        case ShapeOptions.LineEnd.ArrowStealthEnd:
                            this._writer.WriteAttributeString("type", "stealth");
                            break;
                        case ShapeOptions.LineEnd.ArrowDiamondEnd:
                            this._writer.WriteAttributeString("type", "diamond");
                            break;
                        case ShapeOptions.LineEnd.ArrowOvalEnd:
                            this._writer.WriteAttributeString("type", "oval");
                            break;
                        case ShapeOptions.LineEnd.ArrowOpenEnd:
                            this._writer.WriteAttributeString("type", "arrow");
                            break;
                        case ShapeOptions.LineEnd.ArrowChevronEnd: //this should be ignored
                        case ShapeOptions.LineEnd.ArrowDoubleChevronEnd:
                            this._writer.WriteAttributeString("type", "triangle");
                            break;
                    }
                    this._writer.WriteAttributeString("w", "med");
                    this._writer.WriteAttributeString("len", "med");
                    this._writer.WriteEndElement(); //headEnd
                }
            }

            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndArrowhead))
            {
                var val = (ShapeOptions.LineEnd)this.so.OptionsByID[ShapeOptions.PropertyId.lineEndArrowhead].op;
                if (val != ShapeOptions.LineEnd.NoEnd)
                {
                    this._writer.WriteStartElement("a", "tailEnd", OpenXmlNamespaces.DrawingML);
                    switch (val)
                    {
                        case ShapeOptions.LineEnd.ArrowEnd:
                            this._writer.WriteAttributeString("type", "triangle");
                            break;
                        case ShapeOptions.LineEnd.ArrowStealthEnd:
                            this._writer.WriteAttributeString("type", "stealth");
                            break;
                        case ShapeOptions.LineEnd.ArrowDiamondEnd:
                            this._writer.WriteAttributeString("type", "diamond");
                            break;
                        case ShapeOptions.LineEnd.ArrowOvalEnd:
                            this._writer.WriteAttributeString("type", "oval");
                            break;
                        case ShapeOptions.LineEnd.ArrowOpenEnd:
                            this._writer.WriteAttributeString("type", "arrow");
                            break;
                        case ShapeOptions.LineEnd.ArrowChevronEnd: //this should be ignored
                        case ShapeOptions.LineEnd.ArrowDoubleChevronEnd:
                            this._writer.WriteAttributeString("type", "triangle");
                            break;
                    }
                    this._writer.WriteAttributeString("w", "med");
                    this._writer.WriteAttributeString("len", "med");
                    this._writer.WriteEndElement(); //tailnd
                }
            }

            if (!this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndCapStyle))
            {
                //    _writer.WriteStartElement("a", "miter", OpenXmlNamespaces.DrawingML);
                //    _writer.WriteAttributeString("lim", "800000");
                //    _writer.WriteEndElement();
            }

            this._writer.WriteEndElement(); //ln

            //shadow
            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.ShadowStyleBooleanProperties))
            {
                var sp = new ShadowStyleBooleanProperties(this.so.OptionsByID[ShapeOptions.PropertyId.ShadowStyleBooleanProperties].op);
                if (sp.fUsefShadow & sp.fShadow)
                {
                    new ShadowMapping(this._ctx, this._writer).Apply(this.so);
                }
            }
        }

        public Point scanEMFPictureForSize(BlipStoreEntry bse)
        {
            var mb = this._ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
            var size = new Point();

            //write the blip
            if (mb != null)
            {
                switch (bse.btWin32)
                {
                    case BlipStoreEntry.BlipType.msoblipEMF:
                    case BlipStoreEntry.BlipType.msoblipWMF:

                        //it's a meta image
                        var metaBlip = (MetafilePictBlip)mb;
                        size = metaBlip.m_ptSize;

                        //meta images can be compressed
                        var decompressed = metaBlip.Decrompress();

                        break;
                }
            }
            
            return size;
        
        }

        private void writeOle(ShapeContainer container, ExOleEmbedContainer oleContainer)
        {
            var sh = container.FirstChildWithType<Shape>();
            var so = container.FirstChildWithType<ShapeOptions>();

            this._writer.WriteStartElement("p", "graphicFrame", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("p", "nvGraphicFramePr", OpenXmlNamespaces.PresentationML);
            
            //string id = WriteCNvPr(--groupcounter, "");
            string id = WriteCNvPr(sh.spid, "");

            this._writer.WriteStartElement("p", "cNvGraphicFramePr", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("a", "graphicFrameLocks", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("noChangeAspect", "1");
            this._writer.WriteEndElement(); //graphicFrameLocks
            this._writer.WriteEndElement(); //cNvGraphicFramePr
            //_writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");

            this._writer.WriteStartElement("p", "nvPr", OpenXmlNamespaces.PresentationML);
            OEPlaceHolderAtom placeholder = null;
            int exObjIdRef = 0;
            CheckClientData(container.FirstChildWithType<ClientData>(), ref placeholder, ref exObjIdRef);
            this._writer.WriteEndElement();

            this._writer.WriteEndElement(); //nvGraphicFramePr   

            var anchor = new Rectangle();
            var clanchor = container.FirstChildWithType<ClientAnchor>();
            if (clanchor == null)
            {
                 var chanchor = container.FirstChildWithType<ChildAnchor>();
                 anchor = new Rectangle(chanchor.Left, chanchor.Top, chanchor.rcgBounds.Width, chanchor.rcgBounds.Height);
                 if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
                 {
                    this._writer.WriteStartElement("p", "xfrm", OpenXmlNamespaces.PresentationML);

                    this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("x", anchor.Left.ToString());
                    this._writer.WriteAttributeString("y", anchor.Top.ToString());
                    this._writer.WriteEndElement();

                    this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("cx", (anchor.Right - anchor.Left).ToString());
                    this._writer.WriteAttributeString("cy", (anchor.Bottom - anchor.Top).ToString());
                    this._writer.WriteEndElement();

                    this._writer.WriteEndElement();
                 }
            }
            else
            {
                anchor = new Rectangle(clanchor.Left, clanchor.Top,clanchor.Right - clanchor.Left,clanchor.Bottom - clanchor.Top);
                if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
                {
                    this._writer.WriteStartElement("p", "xfrm", OpenXmlNamespaces.PresentationML);

                    this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                    this._writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                    this._writer.WriteEndElement();

                    this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                    this._writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                    this._writer.WriteEndElement();

                    this._writer.WriteEndElement();
                }
            }

            this._writer.WriteStartElement("a", "graphic", OpenXmlNamespaces.DrawingML);

            this._writer.WriteStartElement("a", "graphicData", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("uri", "http://schemas.openxmlformats.org/presentationml/2006/ole");

            EmbeddedObjectPart embPart = null;
            embPart = this.parentSlideMapping.targetPart.AddEmbeddedObjectPart(EmbeddedObjectPart.ObjectType.Other);
            embPart.TargetDirectory = "..\\embeddings";
            var outStream = embPart.GetStream();
            if (oleContainer.stgAtom.Instance == 1)
            {
                outStream.Write(oleContainer.stgAtom.DecompressData(), 0, (int)oleContainer.stgAtom.decompressedSize);
            }
            else
            {
                outStream.Write(oleContainer.stgAtom.data, 0, oleContainer.stgAtom.data.Length);
            }

            string rId = embPart.RelIdToString;

            string spid = "_x0000_s" + sh.spid.ToString(); //+ id;
            string name = oleContainer.AllChildrenWithType<CStringAtom>()[0].Text;
            string progId = oleContainer.AllChildrenWithType<CStringAtom>()[1].Text;

            var gr = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
            var bse = (BlipStoreEntry)gr.FirstChildWithType<BlipStoreContainer>().Children[(int)so.OptionsByID[ShapeOptions.PropertyId.Pib].op - 1];

            //VmlPart vmlPart = null;
            //vmlPart = parentSlideMapping.targetPart.AddVmlPart();
            //vmlPart.TargetDirectory = "..\\drawings";
            //System.IO.Stream vmlStream = vmlPart.GetStream();

            var gsr = container.FirstAncestorWithType<GroupContainer>().FirstChildWithType<ShapeContainer>().FirstChildWithType<GroupShapeRecord>();

            var anch = container.FirstChildWithType<ClientAnchor>();
            if (anch == null) anch = container.FirstAncestorWithType<GroupContainer>().FirstChildWithType<ShapeContainer>().FirstChildWithType<ClientAnchor>();
            var chanch = container.FirstChildWithType<ChildAnchor>();

            Rectangle rec;
            if (chanch != null)
            {
                rec = new Rectangle(chanch.Left, chanch.Top, chanch.Right - anch.Left,chanch.Bottom - anch.Top);
            }
            else if (anch != null)
            {
                rec = new Rectangle(anch.Left, anch.Top, anch.Right - anch.Left, anch.Bottom - anch.Top);
            }
            else
            {
                rec = new Rectangle(chanch.Left, chanch.Top, chanch.Right - chanch.Left, chanch.Bottom - chanch.Top);
            }

            //VMLPictureMapping vm = new VMLPictureMapping(vmlPart, _ctx.WriterSettings);
            var size = new Point();
            //vm.Apply(bse, sh, so, rec, _ctx, spid, ref size);
            addVMLEntry(bse, sh, so, rec, spid, ref size);

            this._writer.WriteStartElement("p", "oleObj", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("spid", spid);
            this._writer.WriteAttributeString("name", name);
            this._writer.WriteAttributeString("id",OpenXmlNamespaces.Relationships, rId);
            this._writer.WriteAttributeString("imgW", size.X.ToString());
            this._writer.WriteAttributeString("imgH", size.Y.ToString());
            this._writer.WriteAttributeString("progId", progId);

            this._writer.WriteStartElement("p", "embed", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("followColorScheme", "full");
            this._writer.WriteEndElement(); //embed

            this._writer.WriteEndElement(); //oleObj

            this._writer.WriteEndElement(); //graphicData

            this._writer.WriteEndElement(); //graphic

            this._writer.WriteEndElement(); //graphicFrame
        }

        private List<ArrayList> VMLEntries = new List<ArrayList>();
        private void addVMLEntry(BlipStoreEntry bse, Shape shape, ShapeOptions options, Rectangle bounds, string spid, ref Point size)
        {
            size = scanEMFPictureForSize(bse);
            var newVMLEntries = new ArrayList();
            newVMLEntries.Add(bse);
            newVMLEntries.Add(shape);
            newVMLEntries.Add(options);
            newVMLEntries.Add(bounds);
            newVMLEntries.Add(spid);
            newVMLEntries.Add(size);
            this.VMLEntries.Add(newVMLEntries);
        }

        private void writeVML()
        {
            if (this.VMLEntries.Count > 0)
            {
                VmlPart vmlPart = null;
                vmlPart = this.parentSlideMapping.targetPart.AddVmlPart();
                vmlPart.TargetDirectory = "..\\drawings";
                var vmlStream = vmlPart.GetStream();

                var vm = new VMLPictureMapping(vmlPart, this._ctx.WriterSettings);
                //var size = new Point();

                vm.Apply(this.VMLEntries, this._ctx);
                //vm.Apply(bse, sh, so, rec, _ctx, spid, ref size);
            }
        }

        private void writePic(ShapeContainer container)
        {
            var sh = container.FirstChildWithType<Shape>();
            var so = container.FirstChildWithType<ShapeOptions>();

            uint indexOfPicture = 0;
            //TODO: read these infos from so
            ++this._ctx.lastImageID;
            int id = this._ctx.lastImageID;
            string name = "";
            string descr = "";
            string rId = "";
            foreach (var en in so.Options)
            {
                switch (en.pid)
                {
                    case ShapeOptions.PropertyId.Pib:
                        indexOfPicture = en.op - 1;
                        break;
                    case ShapeOptions.PropertyId.pibName:
                    //    name = Encoding.Unicode.GetString(en.opComplex);
                    //    name = name.Substring(0, name.Length - 1).Replace("\0","");
                        break;
                    case ShapeOptions.PropertyId.pibPrintName:
                        id = (int)en.op;
                        break;

                }
            }

            var gr = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
            var bse = (BlipStoreEntry)gr.FirstChildWithType<BlipStoreContainer>().Children[(int)indexOfPicture];

            //if (this.parentSlideMapping is MasterMapping) return;
            
            if (this._ctx.AddedImages.ContainsKey(bse.foDelay))
            {
                rId = this._ctx.AddedImages[bse.foDelay]; 
            } else {

                if (!this._ctx.Ppt.PicturesContainer._pictures.ContainsKey(bse.foDelay))
                {
                    return;
                }

                var recBlip = this._ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
                if (recBlip is BitmapBlip)
                {
                    var b = (BitmapBlip)this._ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

                    ImagePart imgPart = null;
                    imgPart = this.parentSlideMapping.targetPart.AddImagePart(getImageType(b.TypeCode));
                    imgPart.TargetDirectory = "..\\media";
                    var outStream = imgPart.GetStream();
                    outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);

                    rId = imgPart.RelIdToString;
                }
                else if (recBlip is MetafilePictBlip)
                {
                    var mb = (MetafilePictBlip)this._ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

                    ImagePart imgPart = null;
                    imgPart = this.parentSlideMapping.targetPart.AddImagePart(getImageType(mb.TypeCode));
                    imgPart.TargetDirectory = "..\\media";
                    var outStream = imgPart.GetStream();

                    var decompressed = mb.Decrompress();
                    outStream.Write(decompressed, 0, decompressed.Length);
                    //outStream.Write(mb.m_pvBits, 0, mb.m_pvBits.Length);

                    rId = imgPart.RelIdToString;
                }
                //this._ctx.AddedImages.Add(bse.foDelay, rId);
            }

            this._writer.WriteStartElement("p", "pic", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "nvPicPr", OpenXmlNamespaces.PresentationML);


            if (!this.spidToId.ContainsKey(sh.spid))
            {
                this.spidToId.Add(sh.spid, id);
            }

            this._writer.WriteStartElement("p", "cNvPr", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", this.spidToId[sh.spid].ToString());

            //_writer.WriteStartElement("p", "cNvPr", OpenXmlNamespaces.PresentationML);
            //_writer.WriteAttributeString("id", id);
            this._writer.WriteAttributeString("name", name);
            this._writer.WriteAttributeString("descr", descr);
            this._writer.WriteEndElement(); //p:cNvPr

            this._writer.WriteStartElement("p", "cNvPicPr", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("a", "picLocks", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("noChangeAspect", "1");
            this._writer.WriteAttributeString("noChangeArrowheads", "1");
            this._writer.WriteEndElement(); //a:picLocks
            this._writer.WriteEndElement(); //p:cNvPicPr

            this._writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");

            this._writer.WriteEndElement(); //p:nvPicPr

            this._writer.WriteStartElement("p", "blipFill", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("embed", OpenXmlNamespaces.Relationships, rId);

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureTransparent))
            {
                RegularContainer slide = so.FirstAncestorWithType<Slide>();
                if (slide == null) slide = so.FirstAncestorWithType<Note>();
                if (slide == null) slide = so.FirstAncestorWithType<Handout>();
                string colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.pictureTransparent].op, slide, so);
                this._writer.WriteStartElement("a", "clrChange", OpenXmlNamespaces.DrawingML);
                this._writer.WriteStartElement("a", "clrFrom", OpenXmlNamespaces.DrawingML);
                this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("val", colorval);
                this._writer.WriteEndElement(); //srgbClr
                this._writer.WriteEndElement(); //clrFrom
                this._writer.WriteStartElement("a", "clrTo", OpenXmlNamespaces.DrawingML);
                this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("val", colorval);
                this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("val", "0");
                this._writer.WriteEndElement(); //alpha
                this._writer.WriteEndElement(); //srgbClr
                this._writer.WriteEndElement(); //clrTo
                this._writer.WriteEndElement(); //clrChange
            }

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureBrightness) | so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureContrast))
            {
                this._writer.WriteStartElement("a", "lum", OpenXmlNamespaces.DrawingML);

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureBrightness))
                {
                    uint b = so.OptionsByID[ShapeOptions.PropertyId.pictureBrightness].op;
                    if (b == 0xFFF8000) b = 0;

                    Decimal b1 = 0;

                    if (((int)b) < 0)
                    {
                        int b2 = (int)b;
                        b1 = (Decimal)b2 / 0x8000;
                    }
                    else
                    {
                        b1 = (Decimal)b / 0x8000;
                    }
                    
                    b1 = b1 * 100000;
                    b1 = Math.Floor(b1);
                    this._writer.WriteAttributeString("bright", b1.ToString());
                }

                //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pictureContrast))
                //{
                //    Int32 b = (int)so.OptionsByID[ShapeOptions.PropertyId.pictureContrast].op;
                //    Decimal b2 = (Decimal)b  / 0x10000; //This comes from analysing, no hint in spec found
                //    b2 = b2 * 100000;
                //    b2 = Math.Floor(b2);
                //    if (b == 0x7FFFFFFF) b2 = 100000;
                //    if (b2 > 100000) b2 = 100000;

                //    _writer.WriteAttributeString("contrast", b2.ToString());
                //}

                this._writer.WriteEndElement();
            }


            this._writer.WriteEndElement(); //a:blip
            this._writer.WriteStartElement("a", "srcRect", OpenXmlNamespaces.DrawingML);
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromLeft))
            {
                this._writer.WriteAttributeString("l", Math.Floor((Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.cropFromLeft].op / 65536 * 100000).ToString());
            }
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromTop))
            {
                this._writer.WriteAttributeString("t", Math.Floor((Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.cropFromTop].op / 65536 * 100000).ToString());
            }
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromRight))
            {
                this._writer.WriteAttributeString("r", Math.Floor((Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.cropFromRight].op / 65536 * 100000).ToString());
            }
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.cropFromBottom))
            {
                this._writer.WriteAttributeString("b", Math.Floor((Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.cropFromBottom].op / 65536 * 100000).ToString());
            }
            this._writer.WriteEndElement();

            this._writer.WriteStartElement("a", "stretch", OpenXmlNamespaces.DrawingML);
            this._writer.WriteElementString("a", "fillRect", OpenXmlNamespaces.DrawingML, "");
            this._writer.WriteEndElement(); //a:stretch
            this._writer.WriteEndElement(); //p:blipFill

            // Visible shape properties
            this._writer.WriteStartElement("p", "spPr", OpenXmlNamespaces.PresentationML);
            var anchor = container.FirstChildWithType<ClientAnchor>();
            if (anchor != null && anchor.Right >= anchor.Left && anchor.Bottom >= anchor.Top)
            {
                this._writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);

                this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("x", Utils.MasterCoordToEMU(anchor.Left).ToString());
                this._writer.WriteAttributeString("y", Utils.MasterCoordToEMU(anchor.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("cx", Utils.MasterCoordToEMU(anchor.Right - anchor.Left).ToString());
                this._writer.WriteAttributeString("cy", Utils.MasterCoordToEMU(anchor.Bottom - anchor.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteEndElement();
            }
            else
            {
                var chanchor = container.FirstChildWithType<ChildAnchor>();

                this._writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);

                this._writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("x", (chanchor.Left).ToString());
                this._writer.WriteAttributeString("y", (chanchor.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("cx", (chanchor.Right - chanchor.Left).ToString());
                this._writer.WriteAttributeString("cy", (chanchor.Bottom - chanchor.Top).ToString());
                this._writer.WriteEndElement();

                this._writer.WriteEndElement();
            }

            this._writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("prst", "rect");
            this._writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
            this._writer.WriteEndElement(); //a:prstGeom
            this._writer.WriteElementString("a", "noFill", OpenXmlNamespaces.DrawingML, "");

            //line
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineStyleBooleans))
            {
                this._writer.WriteStartElement("a", "ln", OpenXmlNamespaces.DrawingML);
                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineWidth))
                {
                    this._writer.WriteAttributeString("w", so.OptionsByID[ShapeOptions.PropertyId.lineWidth].op.ToString());
                }
                WriteLineProperties(so);
                this._writer.WriteEndElement();
            }

            //shadow
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.ShadowStyleBooleanProperties))
            {
                var p = new ShadowStyleBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.ShadowStyleBooleanProperties].op);
                if (p.fUsefShadow & p.fShadow)
                {
                    new ShadowMapping(this._ctx, this._writer).Apply(so);
                }
            }

            this._writer.WriteEndElement(); //p:spPr

            this._writer.WriteEndElement(); //p:pic

        }

        public void writeBodyPr(Record rec, bool cancelAttributes, bool no3D)
        {
            this._writer.WriteStartElement("a", "bodyPr", OpenXmlNamespaces.DrawingML);
            //bool cancelAttributes = false;
            if (rec is ShapeContainer)
            {
                var container = (ShapeContainer)rec;
                switch (container.FirstChildWithType<Shape>().Instance)
                {
                    case 0x88: // WordArt 1, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 18, 19, 25, 29, 30
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.GeometryTextBooleanProperties))
                        {
                            var gbp = new GeometryTextBooleanProperties(this.so.OptionsByID[ShapeOptions.PropertyId.GeometryTextBooleanProperties].op);
                            if (gbp.fUsegtextFVertical && gbp.gtextFVertical)
                            {
                                this._writer.WriteAttributeString("vert", "wordArtVert");
                            }
                        }

                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textPlain");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val 50000");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0x8A: // WordArt 20
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textTriangle");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val 50000");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0x90: // WordArt 3
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textArchUp");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val 10800000");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0x98: // WordArt 23
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textCurveUp");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val 40356");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0x9A: // WordArt 26, 28
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textCascadeUp");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val 44444");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0x9C: // WordArt 17
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textWave1");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj1");
                        this._writer.WriteAttributeString("fmla", "val 13005");
                        this._writer.WriteEndElement();
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj2");
                        this._writer.WriteAttributeString("fmla", "val 0");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0x9E: //WordArt 22
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textDoubleWave1");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj1");
                        this._writer.WriteAttributeString("fmla", "val 6500");
                        this._writer.WriteEndElement();
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj2");
                        this._writer.WriteAttributeString("fmla", "val 0");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0x9F: // WordArt 24
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.GeometryTextBooleanProperties))
                        {
                            var gbp = new GeometryTextBooleanProperties(this.so.OptionsByID[ShapeOptions.PropertyId.GeometryTextBooleanProperties].op);
                            if (gbp.fUsegtextFVertical && gbp.gtextFVertical)
                            {
                                this._writer.WriteAttributeString("vert", "wordArtVert");
                            }
                        }

                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textWave4");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj1");
                        this._writer.WriteAttributeString("fmla", "val 13005");
                        this._writer.WriteEndElement();
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj2");
                        this._writer.WriteAttributeString("fmla", "val 0");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0xA1: // WordArt 4
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textDeflate");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val 26227");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0xA3: // WordArt 27
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textDeflateBottom");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val 76472");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0xAA: //WordArt 21
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textFadeUp");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val 9991");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0xAC: // WordArt 2, 14
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textSlantUp");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val 55556");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    case 0xAF: // WordArt 5
                        this._writer.WriteAttributeString("wrap", "none");
                        this._writer.WriteAttributeString("fromWordArt", "1");

                        this._writer.WriteStartElement("a", "prstTxWarp", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "textCanDown");
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val 33333");
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        cancelAttributes = true;
                        break;
                    default:
                        break;
                }
            }

            if (!cancelAttributes)
            {

                if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dxTextLeft))
                {
                    this._writer.WriteAttributeString("lIns", this.so.OptionsByID[ShapeOptions.PropertyId.dxTextLeft].op.ToString());
                }

                if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dyTextTop))
                {
                    this._writer.WriteAttributeString("tIns", this.so.OptionsByID[ShapeOptions.PropertyId.dyTextTop].op.ToString());
                }

                if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dxTextRight))
                {
                    this._writer.WriteAttributeString("rIns", this.so.OptionsByID[ShapeOptions.PropertyId.dxTextRight].op.ToString());
                }

                if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.dyTextBottom))
                {
                    this._writer.WriteAttributeString("bIns", this.so.OptionsByID[ShapeOptions.PropertyId.dyTextBottom].op.ToString());
                }


                if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.WrapText))
                {
                    switch (this.so.OptionsByID[ShapeOptions.PropertyId.WrapText].op)
                    {
                        case 0: //square
                            this._writer.WriteAttributeString("wrap", "square");
                            break;
                        case 1: //by points
                            break; //TODO
                        case 2: //none
                            this._writer.WriteAttributeString("wrap", "none");
                            break;
                        case 3: //top bottom
                        case 4: //through
                        default:
                            break; //TODO
                    }
                }

                string s = "";
                foreach (var en in this.so.Options)
                {
                    switch (en.pid)
                    {
                        case ShapeOptions.PropertyId.anchorText:

                            switch (en.op)
                            {
                                case 0: //Top
                                    this._writer.WriteAttributeString("anchor", "t");
                                    break;
                                case 1: //Middle
                                    this._writer.WriteAttributeString("anchor", "ctr");
                                    break;
                                case 2: //Bottom
                                    this._writer.WriteAttributeString("anchor", "b");
                                    break;
                                case 3: //TopCentered
                                    this._writer.WriteAttributeString("anchor", "t");
                                    this._writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                                case 4: //MiddleCentered
                                    this._writer.WriteAttributeString("anchor", "ctr");
                                    this._writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                                case 5: //BottomCentered
                                    this._writer.WriteAttributeString("anchor", "b");
                                    this._writer.WriteAttributeString("anchorCtr", "1");
                                    break;
                                case 6: //TopBaseline
                                case 7: //BottomBaseline
                                case 8: //TopCenteredBaseline
                                case 9: //BottomCenteredBaseline
                                    //TODO
                                    break;
                            }
                            break;
                        default:
                            s += en.pid.ToString() + " ";
                            break;
                    }
                }
            }

            RegularContainer slide = this.so.FirstAncestorWithType<Slide>();
            if (slide == null) slide = this.so.FirstAncestorWithType<Note>();
            if (slide == null) slide = this.so.FirstAncestorWithType<Handout>();
            if (!no3D && this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.ThreeDObjectBooleanProperties))
            {
                var tdo = new ThreeDObjectProperties(this.so.OptionsByID[ShapeOptions.PropertyId.ThreeDObjectBooleanProperties].op);

                if (tdo.fc3D && tdo.fUsefc3D)
                {
                    ThreeDStyleProperties tds = null;
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.ThreeDStyleBooleanProperties))
                    {
                        tds = new ThreeDStyleProperties(this.so.OptionsByID[ShapeOptions.PropertyId.ThreeDStyleBooleanProperties].op);
                    }

                    double x = -1;
                    double y = -1;
                    double ox = -1;
                    double oy = -1;
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.c3DOriginX))
                    {
                        var data = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.c3DOriginX].op);
                        ox = new FixedPointNumber(BitConverter.ToUInt16(data, 0), BitConverter.ToUInt16(data, 2)).Value;
                    }
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.c3DOriginY))
                    {
                        var data = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.c3DOriginY].op);
                        oy = new FixedPointNumber(BitConverter.ToUInt16(data, 0), BitConverter.ToUInt16(data, 2)).Value;
                    }
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.c3DKeyX))
                    {
                        var data = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.c3DKeyX].op);
                        x = new FixedPointNumber(BitConverter.ToUInt16(data, 0), BitConverter.ToUInt16(data, 2)).Value;
                    }
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.c3DKeyY))
                    {
                        var data = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.c3DKeyY].op);
                        y = new FixedPointNumber(BitConverter.ToUInt16(data, 0), BitConverter.ToUInt16(data, 2)).Value;
                    }
                    string prst = "legacyObliqueRight";

                    if (ox == -1 && oy == 0)
                    {
                        prst = "legacyObliqueRight";
                    }
                    else if (((int)ox) == 32768 && oy == -1)
                    {
                        prst = "legacyPerspectiveTopLeft";
                    }
                    else if (ox == 0 && oy == 0)
                    {
                        prst = "legacyPerspectiveFront";
                    }
                    else if (ox == -1 && ((int)oy) == 32768)
                    {
                        prst = "legacyPerspectiveBottomRight";
                    }

                    string dir = "t";
                    if (((int)x) == 15536)
                    {
                        dir = "t";
                    }
                    else if (((int)y) == 15536)
                    {
                        dir = "r";
                    }
                    else if (x == -1)
                    {
                        dir = "b";
                    }
                    else if (x == 0 && y == 50000)
                    {
                        dir = "l";
                    }

                    string rig = "legacyHarsh3";
                    if (ox == 0 && oy == 0 && ((int)x) == 15536 && ((int)y) == 15536)
                    {
                        rig = "legacyNormal2";
                    }
                    else if ((((int)ox) == 32768) || (ox == 0 && oy == 0))
                    {
                        rig = "legacyNormal3";
                    }


                    this._writer.WriteStartElement("a", "scene3d", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteStartElement("a", "camera", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("prst", prst);

                    if (tds != null)
                        if (tds.fc3DConstrainRotation && tds.fUsefc3DConstrainRotation)
                        {
                            decimal xra = 0;
                            decimal yra = 0;
                            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.c3DXRotationAngle))
                            {
                                var data = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.c3DXRotationAngle].op);
                                int integral = BitConverter.ToInt16(data, 2);
                                uint fractional = BitConverter.ToUInt16(data, 0);
                                xra = -1 * (integral + ((decimal)fractional / (decimal)65536));
                                if (xra < 0) xra += 360;
                            }
                            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.c3DYRotationAngle))
                            {
                                var data = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.c3DYRotationAngle].op);
                                int integral = BitConverter.ToInt16(data, 2);
                                uint fractional = BitConverter.ToUInt16(data, 0);
                                yra = integral + ((decimal)fractional / (decimal)65536);
                                if (yra < 0) yra += 360;
                            }

                            if (xra != 0 || yra != 0)
                            {
                                //rot
                                this._writer.WriteStartElement("a", "rot", OpenXmlNamespaces.DrawingML);
                                //@lat
                                this._writer.WriteAttributeString("lat", Math.Floor(xra * 60000).ToString());
                                //@lon
                                this._writer.WriteAttributeString("lon", Math.Floor(yra * 60000).ToString());
                                //@rev
                                this._writer.WriteAttributeString("rev", "0");
                                this._writer.WriteEndElement(); //rot
                            }
                        }

                    this._writer.WriteEndElement(); //camera
                    this._writer.WriteStartElement("a", "lightRig", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("rig", rig);
                    this._writer.WriteAttributeString("dir", dir);
                    this._writer.WriteEndElement(); //lightRig
                    this._writer.WriteEndElement(); //scene3d

                    this._writer.WriteStartElement("a", "sp3d", OpenXmlNamespaces.DrawingML);

                    string extrusionH = "457200";
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.c3DExtrudeBackward))
                    {
                        //the -27000 comes from analysing PP 2003
                        uint v = this.so.OptionsByID[ShapeOptions.PropertyId.c3DExtrudeBackward].op - 27000;
                        extrusionH = v.ToString();
                    }

                    this._writer.WriteAttributeString("extrusionH", extrusionH);
                    this._writer.WriteAttributeString("prstMaterial", "legacyMatte");

                    //if (tdo.fc3UseExtrusionColor && tdo.fUsefc3DUseExtrusionColor)
                    if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.c3DExtrusionColor))
                    {
                        this._writer.WriteStartElement("a", "extrusionClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", Utils.getRGBColorFromOfficeArtCOLORREF(this.so.OptionsByID[ShapeOptions.PropertyId.c3DExtrusionColor].op, slide, this.so));
                        this._writer.WriteEndElement(); //srgbClr
                        this._writer.WriteEndElement(); //extrusionClr
                    }

                    this._writer.WriteEndElement(); //sp3d
                }
            }

            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.TextBooleanProperties))
            {
                var props = new TextBooleanProperties(this.so.OptionsByID[ShapeOptions.PropertyId.TextBooleanProperties].op);
                if (props.fFitShapeToText && props.fUsefFitShapeToText) this._writer.WriteElementString("a", "spAutoFit", OpenXmlNamespaces.DrawingML, "");
            }

            this._writer.WriteEndElement();
        }

        private void CheckClientData(ClientData clientData, ref OEPlaceHolderAtom placeholder, ref int exObjIdRef)
        {
            bool output = exObjIdRef > -1;
            this.ShapeStyleTextProp9Atom = null;
            bool phWritten = false;
            if (clientData != null)
            {

                var ms = new System.IO.MemoryStream(clientData.bytes);

                if (ms.Length > 0)
                {
                    var rec = Record.ReadRecord(ms);
                    bool blnContinue = true; 
                    if (rec.TypeCode == 4116 && output)
                    {
                        var animinfo = (AnimationInfoContainer)rec;
                        this.animinfos.Add(animinfo, this._idCnt);
                        if (ms.Position < ms.Length)
                        {
                            rec = Record.ReadRecord(ms);
                            rec.SiblingIdx = 1;
                        }
                        else
                        {
                            blnContinue = false;
                        }
                    }

                    if (blnContinue)
                    while (true)
                    {
                        switch (rec.TypeCode)
                        {
                            case 3009:
                                exObjIdRef = ((ExObjRefAtom)rec).exObjIdRef;
                                break;
                            case 3011:
                                placeholder = (OEPlaceHolderAtom)rec;

                                if (placeholder != null && output)
                                {

                                        this._writer.WriteStartElement("p", "ph", OpenXmlNamespaces.PresentationML);

                                    if (!placeholder.IsObjectPlaceholder())
                                    {
                                        string typeValue = Utils.PlaceholderIdToXMLValue(placeholder.PlacementId);
                                            this._writer.WriteAttributeString("type", typeValue);
                                    }

                                    switch (placeholder.PlaceholderSize)
                                    {
                                        case 1:
                                                this._writer.WriteAttributeString("sz", "half");
                                            break;
                                        case 2:
                                                this._writer.WriteAttributeString("sz", "quarter");
                                            break;
                                    }


                                    if (placeholder.Position != -1)
                                    {
                                            this._writer.WriteAttributeString("idx", placeholder.Position.ToString());
                                    }
                                    else
                                    {
                                        try
                                        {
                                            var master = this._ctx.Ppt.FindMasterRecordById(clientData.FirstAncestorWithType<Slide>().FirstChildWithType<SlideAtom>().MasterId);
                                            foreach (var cont in master.FirstChildWithType<PPDrawing>().FirstChildWithType<DrawingContainer>().FirstChildWithType<GroupContainer>().AllChildrenWithType<ShapeContainer>())
                                            {
                                                var s = cont.FirstChildWithType<Shape>();
                                                var d = cont.FirstChildWithType<ClientData>();
                                                if (d != null)
                                                {
                                                    ms = new System.IO.MemoryStream(d.bytes);
                                                    rec = Record.ReadRecord(ms);
                                                    if (rec is OEPlaceHolderAtom)
                                                    {
                                                        var placeholder2 = (OEPlaceHolderAtom)rec;
                                                        if (placeholder2.PlacementId == PlaceholderEnum.MasterBody && (placeholder.PlacementId == PlaceholderEnum.Body || placeholder.PlacementId == PlaceholderEnum.Object))
                                                        {
                                                            if (placeholder2.Position != -1)
                                                            {
                                                                    this._writer.WriteAttributeString("idx", placeholder2.Position.ToString());
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                        }

                                        catch (Exception)
                                        {
                                            //ignore
                                        }
                                    }

                                        this._writer.WriteEndElement();
                                    phWritten = true;
                                }
                                break;
                            case 4116:
                                var animinfo = (AnimationInfoContainer)rec;
                                    this.animinfos.Add(animinfo, this._idCnt);
                                break;
                            case 5000:
                                var con = (RegularContainer)rec;
                                foreach (var t in con.AllChildrenWithType<ProgBinaryTag>())
                                {
                                    var c = t.FirstChildWithType<CStringAtom>();
                                    var b = t.FirstChildWithType<ProgBinaryTagDataBlob>();
                                    var p = b.FirstChildWithType<StyleTextProp9Atom>();
                                        this.ShapeStyleTextProp9Atom = p;
                                }
                                break;
                            default:
                                break;
                        }
                        if (ms.Position < ms.Length)
                        {
                            rec = Record.ReadRecord(ms);
                        }
                        else
                        {
                            break;
                        }
                    }                    
                }
            
                var container = (RegularContainer)(clientData.ParentRecord);
                foreach (var b in container.AllChildrenWithType<ClientTextbox>())
                {
                    ms = new System.IO.MemoryStream(b.Bytes);
                    Record rec;
                    while (ms.Position < ms.Length)
                    {
                        rec = Record.ReadRecord(ms);

                        switch (rec.TypeCode)
                        {
                            case 0xfa0: //TextCharsAtom
                            case 0xfa1: //TextRunStyleAtom
                            case 0xfa6: //TextRulerAtom
                            case 0xfa8: //TextBytesAtom
                            case 0xfaa: //TextSpecialInfoAtom
                            case 0xfa2: //MasterTextPropAtom
                                break;
                            case 0xfd8: //SlideNumberMCAtom
                               
                                break;
                            case 0xff7: //DateTimeMCAtom
                                if (!phWritten && output)
                                {
                                    this._writer.WriteStartElement("p", "ph", OpenXmlNamespaces.PresentationML);
                                    this._writer.WriteAttributeString("type", "dt");
                                    this._writer.WriteEndElement();
                                }
                                break;
                            case 0xff9: //HeaderMCAtom
                                break;
                            case 0xffa: //FooterMCAtom
                                break;
                            case 0xff8: //GenericDateMCAtom
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
        }

        public void Apply(ClientTextbox textbox)
        {
            Apply(textbox, false);
        }

        public void Apply(ClientTextbox textbox, bool insideTable)
        {
            if (insideTable)
            {
                this._writer.WriteStartElement("a", "txBody", OpenXmlNamespaces.DrawingML);
            }
            else
            {
                this._writer.WriteStartElement("p", "txBody", OpenXmlNamespaces.PresentationML);
            }
            writeBodyPr(textbox, insideTable, true);

            this._writer.WriteStartElement("a", "lstStyle", OpenXmlNamespaces.DrawingML);
            bool lvlRprWritten = false;

            var ms = new System.IO.MemoryStream(textbox.Bytes);
            TextHeaderAtom thAtom = null;
            TextStyleAtom style = null;
            var lst = new List<int>();
            string lang = "";
            string altLang = "";
            while (ms.Position < ms.Length)
            {
                var rec = Record.ReadRecord(ms);

                switch (rec.TypeCode)
                {
                    case 0xf9e: //OutlineTextRefAtom
                        var otrAtom = (OutlineTextRefAtom)rec;
                        var slideListWithText = this._ctx.Ppt.DocumentRecord.RegularSlideListWithText;

                        var thAtoms = slideListWithText.SlideToPlaceholderTextHeaders[textbox.FirstAncestorWithType<Slide>().PersistAtom];
                        thAtom = thAtoms[otrAtom.Index];

                        //if (thAtom.TextAtom != null) text = thAtom.TextAtom.Text;
                        if (thAtom.TextStyleAtom != null) style = thAtom.TextStyleAtom;

                        break;
                    case 0xf9f: //TextHeaderAtom
                        thAtom = (TextHeaderAtom)rec;
                        break;
                    case 0xfa0: //TextCharsAtom
                        thAtom.TextAtom = (TextAtom)rec;
                        break;
                    case 0xfa1: //StyleTextPropAtom
                        style = (TextStyleAtom)rec;
                        style.TextHeaderAtom = thAtom;
                        break;
                    case 0xfa2: //MasterTextPropAtom
                        var m = (MasterTextPropAtom)rec;
                        foreach(var r in m.MasterTextPropRuns)
                        {
                            if (!lst.Contains(r.indentLevel))
                            {
                                this._writer.WriteStartElement("a", "lvl" + (r.indentLevel + 1) + "pPr", OpenXmlNamespaces.DrawingML);

                                if (thAtom.TextType == TextType.CenterTitle || thAtom.TextType == TextType.CenterBody)
                                {
                                    this._writer.WriteAttributeString("algn", "ctr");
                                }

                                //_writer.WriteElementString("a", "buNone", OpenXmlNamespaces.DrawingML, "");

                                this._writer.WriteEndElement();
                                lst.Add(r.indentLevel);
                            }
                        }
                        break;
                    case 0xfa8: //TextBytesAtom
                        //text = ((TextBytesAtom)rec).Text;
                        thAtom.TextAtom = (TextAtom)rec;
                        break;
                    case 0xfaa: //TextSpecialInfoAtom
                        var sia = (TextSpecialInfoAtom)rec;
                        if (sia.Runs.Count > 0)
                        {
                            if (sia.Runs[0].si.lang)
                            {
                                switch (sia.Runs[0].si.lid)
                                {
                                    case 0x0: // no language
                                        break;
                                    case 0x13: //Any Dutch language is preferred over non-Dutch languages when proofing the text
                                        break;
                                    case 0x400: //no proofing
                                        break;
                                    default:
                                        try
                                        {
                                            lang = System.Globalization.CultureInfo.GetCultureInfo(sia.Runs[0].si.lid).IetfLanguageTag;
                                        }
                                        catch (Exception)
                                        {   
                                            //ignore
                                        }                                       
                                        break;
                                }                               
                            }
                            if (sia.Runs[0].si.altLang)
                            {
                                switch (sia.Runs[0].si.altLid)
                                {
                                    case 0x0: // no language
                                        break;
                                    case 0x13: //Any Dutch language is preferred over non-Dutch languages when proofing the text
                                        break;
                                    case 0x400: //no proofing
                                        break;
                                    default:
                                        try
                                        {
                                            altLang = System.Globalization.CultureInfo.GetCultureInfo(sia.Runs[0].si.altLid).IetfLanguageTag;
                                        }
                                        catch (Exception)
                                        {
                                            //ignore
                                        }
                                        break;
                                }
                            }
                        }
                        break;
                    case 0xfd8: //SlideNumberMCAtom
                    case 0xff9: //HeaderMCAtom
                    case 0xffa: //FooterMCAtom
                    case 0xff8: //GenericDateMCAtom
                        if (!lvlRprWritten)
                        foreach (var r in style.PRuns)
                        {
                                this._writer.WriteStartElement("a", "lvl" + (r.IndentLevel + 1) + "pPr", OpenXmlNamespaces.DrawingML);
                            if (r.AlignmentPresent)
                            {
                                switch (r.Alignment)
                                {
                                    case 0x0000: //Left
                                            this._writer.WriteAttributeString("algn", "l");
                                        break;
                                    case 0x0001: //Center
                                            this._writer.WriteAttributeString("algn", "ctr");
                                        break;
                                    case 0x0002: //Right
                                            this._writer.WriteAttributeString("algn", "r");
                                        break;
                                    case 0x0003: //Justify
                                            this._writer.WriteAttributeString("algn", "just");
                                        break;
                                    case 0x0004: //Distributed
                                            this._writer.WriteAttributeString("algn", "dist");
                                        break;
                                    case 0x0005: //ThaiDistributed
                                            this._writer.WriteAttributeString("algn", "thaiDist");
                                        break;
                                    case 0x0006: //JustifyLow
                                            this._writer.WriteAttributeString("algn", "justLow");
                                        break;
                                }
                            }
                            string lastColor = "";
                            string lastSize = "";
                            string lastTypeface = "";
                            RegularContainer slide = textbox.FirstAncestorWithType<Slide>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Note>();
                            if (slide == null) slide = textbox.FirstAncestorWithType<Handout>();

                            new CharacterRunPropsMapping(this._ctx, this._writer).Apply(style.CRuns[0], "defRPr", slide, ref lastColor, ref lastSize, ref lastTypeface, lang, altLang, null,r.IndentLevel,null,null,0, insideTable);
                                this._writer.WriteEndElement();
                            lvlRprWritten = true;
                        }
                        break;
                    default:
                        break;
                }
            }

            this._writer.WriteEndElement();

            new TextMapping(this._ctx, this._writer).Apply(this, textbox, this._footertext, this._headertext, this._datetext, insideTable);

            this._writer.WriteEndElement();
        }

        public static ImagePart.ImageType getImageType(uint TypeCode)
        {
            switch (TypeCode)
            {
                case 0xF01A:
                    return ImagePart.ImageType.Emf;
                case 0xF01B:
                    return ImagePart.ImageType.Wmf;
                case 0xF01D:
                    return ImagePart.ImageType.Jpeg;
                case 0xF01E:
                    return ImagePart.ImageType.Png;
                case 0xF01F: //DIP
                    return ImagePart.ImageType.Bmp;
                case 0xF020:
                    return ImagePart.ImageType.Tiff;
                default:
                    return ImagePart.ImageType.Png;
            }
        }


        public void Apply(RegularContainer container)
        {
            // Descend into container records by default
            foreach (var record in container.Children)
            {
                DynamicApply(record);
            }
        }

        public void Apply(Record record)
        {
            // Ignore unsupported records
            //TraceLogger.DebugInternal("Unsupported record: {0}", record);
        }

        private void WriteGroupShapeProperties(ShapeContainer header)
        {
            var groupShape = header.FirstChildWithType<GroupShapeRecord>();

            // Write non-visible Group Shape properties
            this._writer.WriteStartElement("p", "nvGrpSpPr", OpenXmlNamespaces.PresentationML);

            // Non-visible Canvas Properties
            WriteCNvPr(-1, "");

            this._writer.WriteElementString("p", "cNvGrpSpPr", OpenXmlNamespaces.PresentationML, "");
            this._writer.WriteElementString("p", "nvPr", OpenXmlNamespaces.PresentationML, "");

            this._writer.WriteEndElement();


            // Write visible Group Shape properties
            this._writer.WriteStartElement("p", "grpSpPr", OpenXmlNamespaces.PresentationML);
            WriteXFrm(this._writer, new Rectangle()); // groupShape.rcgBounds

            this._writer.WriteEndElement();
        }

        private void WritecustGeom(Shape sh)
        {

            if (!this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pVertices) | !this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pSegmentInfo))
            {
                this._writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("prst", "rect");
                this._writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                this._writer.WriteEndElement(); //prstGeom
                return;
            }

            this._writer.WriteStartElement("a", "custGeom", OpenXmlNamespaces.DrawingML);

            this._writer.WriteStartElement("a", "cxnLst", OpenXmlNamespaces.DrawingML);

            var pVertices = this.so.OptionsByID[ShapeOptions.PropertyId.pVertices];
            
            uint shapepath = 1;
            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shapePath))
            {
                var ShapePath = this.so.OptionsByID[ShapeOptions.PropertyId.shapePath];
                shapepath = ShapePath.op;
            }
            else
            {
                //shapepath = 4; //complex
            }
            var SegementInfo = this.so.OptionsByID[ShapeOptions.PropertyId.pSegmentInfo];


            PathParser pp;
            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.pGuides))
            {
                var pGuides = this.so.OptionsByID[ShapeOptions.PropertyId.pGuides];
                pp = new PathParser(SegementInfo.opComplex, pVertices.opComplex, pGuides.opComplex);
            } else 
            {
                pp = new PathParser(SegementInfo.opComplex, pVertices.opComplex);
            }
        

            

            foreach (var point in pp.Values)
            {
                this._writer.WriteStartElement("a", "cxn", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("ang", "0");
                this._writer.WriteStartElement("a", "pos", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("x", point.X.ToString());
                this._writer.WriteAttributeString("y", point.Y.ToString());
                this._writer.WriteEndElement(); //pos
                this._writer.WriteEndElement(); //cxn
            }
            this._writer.WriteEndElement(); //cxnLst

            this._writer.WriteStartElement("a", "rect", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("l", "0");
            this._writer.WriteAttributeString("t", "0");
            this._writer.WriteAttributeString("r", "r");
            this._writer.WriteAttributeString("b", "b");
            this._writer.WriteEndElement(); //rect

            this._writer.WriteStartElement("a", "pathLst", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "path", OpenXmlNamespaces.DrawingML);
            //compute width and height:
            int minX = 1000;
            int minY = 1000;
            int maxX = -1000;
            int maxY = -1000;
            foreach (var p in pp.Values)
            {
                if (p.X > 0)
                {
                    if ((p.X) < minX) minX = p.X;
                    if ((p.X) > maxX) maxX = p.X;
                }
                if (p.Y > 0)
                {
                    if ((p.Y) < minY) minY = p.Y;
                    if ((p.Y) > maxY) maxY = p.Y;
                }
            }

            int w = maxX - minX;
            if (w < 0) w = 0;

            int h = maxY - minY;
            if (h < 0) h = 0;

            this._writer.WriteAttributeString("w", w.ToString());
            this._writer.WriteAttributeString("h", h.ToString());

            int valuePointer = 0;

            switch (shapepath)
            {
                case 0: //lines
                case 1: //lines closed
                    while (valuePointer < pp.Values.Count)
                    {
                        if (valuePointer == 0)
                        {
                            this._writer.WriteStartElement("a", "moveTo", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                            this._writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                            this._writer.WriteEndElement(); //pr
                            this._writer.WriteEndElement(); //moveTo
                            valuePointer += 1;
                        }
                        else
                        {
                            this._writer.WriteStartElement("a", "lnTo", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                            this._writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                            this._writer.WriteEndElement(); //pt
                            this._writer.WriteEndElement(); //lnTo
                            valuePointer += 1;
                        }
                    }
                    break;
                case 2: //curves
                    break;
                case 3: //curves closed
                    break;
                case 4: //complex

                    var Escapes = new Dictionary<Point, int>();
                    var tempEscapes = new List<int>();
                    int i = 0;
                    int start = 0;
                    int end;
                    foreach (var seg in pp.Segments)
                    {
                        if (seg.Type == PathSegment.SegmentType.msopathEscape)
                        {
                            tempEscapes.Add(seg.EscapeCode);
                        }
                        if (seg.Type == PathSegment.SegmentType.msopathEnd)
                        {
                            end = i;
                            foreach (int escape in tempEscapes)
                            {
                                if (!Escapes.ContainsKey(new Point(start, end)))
                                Escapes.Add(new Point(start, end), escape);
                            }
                            start = i + 1;
                            tempEscapes.Clear();
                        }
                        i++;
                    }

                    i = 0;
                    foreach (var seg in pp.Segments)
                    {
                        if (valuePointer >= pp.Values.Count) break;

                        if (i == 0)
                        {
                            //check for escape codes
                            foreach (var p in Escapes.Keys)
                            {
                                if (p.X == 0)
                                {
                                    switch (Escapes[p])
                                    {
                                        case 0xA:
                                            this._writer.WriteAttributeString("fill", "none");
                                            break;
                                        case 0xB:
                                            this._writer.WriteAttributeString("stroke", "0");
                                            break;
                                    }
                                }
                            }
                        }

                        switch (seg.Type)
                        {
                            case PathSegment.SegmentType.msopathLineTo:
                                this._writer.WriteStartElement("a", "lnTo", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                                this._writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                                this._writer.WriteEndElement(); //pt
                                this._writer.WriteEndElement(); //lnTo
                                valuePointer += 1;
                                break;
                            case PathSegment.SegmentType.msopathCurveTo:
                                this._writer.WriteStartElement("a", "cubicBezTo", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                                this._writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                                this._writer.WriteEndElement(); //pt
                                valuePointer += 1;
                                this._writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                                this._writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                                this._writer.WriteEndElement(); //pt
                                valuePointer += 1;
                                this._writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                                this._writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                                this._writer.WriteEndElement(); //pt
                                valuePointer += 1;
                                this._writer.WriteEndElement(); //cubicBezTo
                                break;
                            case PathSegment.SegmentType.msopathMoveTo:
                                this._writer.WriteStartElement("a", "moveTo", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteStartElement("a", "pt", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("x", pp.Values[valuePointer].X.ToString());
                                this._writer.WriteAttributeString("y", pp.Values[valuePointer].Y.ToString());
                                this._writer.WriteEndElement(); //pr
                                this._writer.WriteEndElement(); //moveTo
                                valuePointer += 1;
                                break;
                            case PathSegment.SegmentType.msopathClose:
                                this._writer.WriteElementString("a", "close", OpenXmlNamespaces.DrawingML, "");
                                break;
                            case PathSegment.SegmentType.msopathEnd:
                                this._writer.WriteEndElement(); //path
                                this._writer.WriteStartElement("a", "path", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("w", (maxX - minX).ToString());
                                this._writer.WriteAttributeString("h", (maxY - minY).ToString());

                                //check for escape codes
                                foreach (var p in Escapes.Keys)
                                {
                                    if (p.X <= i+1 && p.Y >= i+1)
                                    {
                                        switch (Escapes[p])
                                        {
                                            case 0xA:
                                                this._writer.WriteAttributeString("fill", "none");
                                                break;
                                            case 0xB:
                                                this._writer.WriteAttributeString("stroke", "0");
                                                break;
                                        }
                                    }
                                }


                                break;
                            default:
                                break;
                        }
                        i++;
                    }
                    break;
            }

            this._writer.WriteEndElement(); //path
            this._writer.WriteEndElement(); //pathLst

            this._writer.WriteEndElement(); //custGeom
        }

        private void WriteprstGeom(Shape shape)
        {
            if (shape != null)
            {
                string prst = Utils.getPrstForShape(shape.Instance);
                if (prst.Length > 0)
                {

                    this._writer.WriteStartElement("a", "prstGeom", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("prst", prst);
                    if (prst == "roundRect" & this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.adjustValue)) //TODO: implement for all shapes
                    {
                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj");
                        this._writer.WriteAttributeString("fmla", "val " + Math.Floor(this.so.OptionsByID[ShapeOptions.PropertyId.adjustValue].op * 4.63).ToString()); //TODO: find out where this 4.63 comes from (value found by analysing behaviour of Powerpoint 2003)
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                    }
                    else if (prst == "leftArrow" & this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndArrowWidth) && this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.lineEndArrowLength))
                    {
                        uint w = this.so.OptionsByID[ShapeOptions.PropertyId.lineEndArrowWidth].op;
                        uint l = this.so.OptionsByID[ShapeOptions.PropertyId.lineEndArrowLength].op;

                        if (w == 2 && l == 2)
                        {
                            this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("name", "adj1");
                            this._writer.WriteAttributeString("fmla", "val 50000");
                            this._writer.WriteEndElement();
                            this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("name", "adj2");
                            this._writer.WriteAttributeString("fmla", "val 210000");
                            this._writer.WriteEndElement();
                            this._writer.WriteEndElement();
                        }

                    }
                    else if ((prst == "wedgeRectCallout" || prst == "cloudCallout" || prst == "wedgeEllipseCallout") && this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.adjustValue))
                    {
                        //the following computations are based on experiments using Powerpoint 2003 and are not part of the spec
                        decimal val = (Decimal)(int)this.so.OptionsByID[ShapeOptions.PropertyId.adjustValue].op;
                        decimal percent = val / 21600 * 100;
                        int newVal = 0;
                        if (percent >= 50)
                        {
                            newVal = (int)(percent - 50) * 1000;
                        }
                        else
                        {
                            newVal = (int)(50 - percent) * -1000;
                        }

                        this._writer.WriteStartElement("a", "avLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("name", "adj1");
                        this._writer.WriteAttributeString("fmla", "val " + newVal.ToString());
                        this._writer.WriteEndElement();
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.adjust2Value))
                        {
                            val = (Decimal)(int)this.so.OptionsByID[ShapeOptions.PropertyId.adjust2Value].op;
                            percent = val / 21600 * 100;
                            newVal = 0;
                            if (percent >= 50)
                            {
                                newVal = (int)(percent - 50) * 1000;
                            }
                            else
                            {
                                newVal = (int)(50 - percent) * -1000;
                            }
                            this._writer.WriteStartElement("a", "gd", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("name", "adj2");
                            this._writer.WriteAttributeString("fmla", "val " + newVal.ToString());
                            this._writer.WriteEndElement();
                        }
                        this._writer.WriteEndElement();

                    }
                    else
                    {
                        this._writer.WriteElementString("a", "avLst", OpenXmlNamespaces.DrawingML, "");
                    }
                    this._writer.WriteEndElement(); //prstGeom
                }                
            }
        }

        

        public Dictionary<int, int> spidToId = new Dictionary<int, int>();
        private string WriteCNvPr(int spid, string name)
        {
            string id = "";
            if (!this.spidToId.ContainsKey(spid))
            {
                this.spidToId.Add(spid, ++this._idCnt);
            }

            this._writer.WriteStartElement("p", "cNvPr", OpenXmlNamespaces.PresentationML);
            id = this.spidToId[spid].ToString();
            this._writer.WriteAttributeString("id", id);
            this._writer.WriteAttributeString("name", name);
            this._writer.WriteEndElement();
            return id;
        }

        private void WriteXFrm(XmlWriter _writer, Rectangle rect)
        {
            _writer.WriteStartElement("a", "xfrm", OpenXmlNamespaces.DrawingML);

            // TODO: Coordinate conversion?
            _writer.WriteStartElement("a", "off", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("x", rect.X.ToString());
            _writer.WriteAttributeString("y", rect.Y.ToString());
            _writer.WriteEndElement();

            // TODO: Coordinate conversion?
            _writer.WriteStartElement("a", "ext", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("cx", rect.Width.ToString());
            _writer.WriteAttributeString("cy", rect.Height.ToString());
            _writer.WriteEndElement();

            // TODO: Where do we get this from?
            _writer.WriteStartElement("a", "chOff", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("x", "0");
            _writer.WriteAttributeString("y", "0");
            _writer.WriteEndElement();

            _writer.WriteStartElement("a", "chExt", OpenXmlNamespaces.DrawingML);
            _writer.WriteAttributeString("cx", rect.Width.ToString());
            _writer.WriteAttributeString("cy", rect.Height.ToString());
            _writer.WriteEndElement();            

            _writer.WriteEndElement();
        }
    }
}
