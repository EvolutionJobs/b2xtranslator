using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class OfficeArtContent
    {
        public enum DrawingType
        {
            MainDocument,
            Header
        }

        public struct OfficeArtWordDrawing
        {
            public DrawingType dgglbl;
            public DrawingContainer container;
        }

        public DrawingGroup DrawingGroupData;
        public List<OfficeArtWordDrawing> Drawings;

        public OfficeArtContent(FileInformationBlock fib, VirtualStream tableStream)
        {
            VirtualStreamReader reader = new VirtualStreamReader(tableStream);
            tableStream.Seek(fib.fcDggInfo, System.IO.SeekOrigin.Begin);

             if (fib.lcbDggInfo > 0)
            {
                int maxPosition = (int)(fib.fcDggInfo + fib.lcbDggInfo);

                //read the DrawingGroupData
                this.DrawingGroupData = (DrawingGroup)Record.ReadRecord(reader);

                //read the Drawings
                this.Drawings = new List<OfficeArtWordDrawing>();
                while (reader.BaseStream.Position < maxPosition)
                {
                    OfficeArtWordDrawing drawing = new OfficeArtWordDrawing();
                    drawing.dgglbl = (DrawingType)reader.ReadByte();
                    drawing.container = (DrawingContainer)Record.ReadRecord(reader);

                    for (int i = 0; i < drawing.container.Children.Count; i++)
                    {
                        Record groupChild = drawing.container.Children[i];
                        if (groupChild.TypeCode == 0xF003)
                        {
                            // the child is a subgroup
                            GroupContainer group = (GroupContainer)drawing.container.Children[i];
                            group.Index = i;
                            drawing.container.Children[i] = group;
                        }
                        else if (groupChild.TypeCode == 0xF004)
                        {
                            // the child is a shape
                            ShapeContainer shape = (ShapeContainer)drawing.container.Children[i];
                            shape.Index = i;
                            drawing.container.Children[i] = shape;
                        }
                    }

                    this.Drawings.Add(drawing);
                }
            }
        }

        /// <summary>
        /// Searches the matching shape
        /// </summary>
        /// <param name="spid">The shape ID</param>
        /// <returns>The ShapeContainer</returns>
        public ShapeContainer GetShapeContainer(int spid)
        {
            ShapeContainer ret = null;

            foreach(OfficeArtWordDrawing drawing in this.Drawings)
            {
                GroupContainer group = (GroupContainer)drawing.container.FirstChildWithType<GroupContainer>();
                if (group != null)
                {
                    for (int i = 1; i < group.Children.Count; i++)
                    {
                        Record groupChild = group.Children[i];
                        if (groupChild.TypeCode == 0xF003)
                        {
                            //It's a group of shapes
                            GroupContainer subgroup = (GroupContainer)groupChild;

                            //the referenced shape must be the first shape in the group
                            ShapeContainer container = (ShapeContainer)subgroup.Children[0];
                            Shape shape = (Shape)container.Children[1];
                            if (shape.spid == spid)
                            {
                                ret = container;
                                break;
                            }
                        }
                        else if (groupChild.TypeCode == 0xF004)
                        {
                            //It's a singe shape
                            ShapeContainer container = (ShapeContainer)groupChild;
                            Shape shape = (Shape)container.Children[0];
                            if (shape.spid == spid)
                            {
                                ret = container;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    continue;
                }

                if (ret != null)
                {
                    break;
                }
            }

            return ret;

        }
    }
}
