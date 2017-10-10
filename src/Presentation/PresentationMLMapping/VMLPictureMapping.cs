using System.Collections.Generic;
using System.Text;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using System.IO;
using System.Drawing;
using b2xtranslator.Tools;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.OfficeDrawing.Shapetypes;
using System.Collections;

namespace b2xtranslator.PresentationMLMapping
{
    public class VMLPictureMapping
        : AbstractOpenXmlMapping
    {
        ContentPart _targetPart;
        private ConversionContext _ctx;

        public VMLPictureMapping(VmlPart vmlPart, XmlWriterSettings xws)
            : base(XmlWriter.Create(vmlPart.GetStream(), xws))
        {
            this._targetPart = vmlPart;
        }
        
        //public void Apply(BlipStoreEntry bse, Shape shape, ShapeOptions options, Rectangle bounds, ConversionContext ctx, string spid, ref Point size)
        public void Apply(List<ArrayList> VMLEntriesList, ConversionContext ctx)
        {
            this._ctx = ctx;
            BlipStoreEntry bse;

            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("xml");

            this._writer.WriteStartElement("o", "shapelayout", OpenXmlNamespaces.Office);
            this._writer.WriteAttributeString("v", "ext", OpenXmlNamespaces.VectorML, "edit");
            this._writer.WriteStartElement("o", "idmap", OpenXmlNamespaces.Office);
            this._writer.WriteAttributeString("v", "ext", OpenXmlNamespaces.VectorML, "edit");
            this._writer.WriteAttributeString("data", "1079");
            this._writer.WriteEndElement(); //idmap
            this._writer.WriteEndElement(); //shapelayout

            //v:shapetype
            var type = new PictureFrameType();
            type.Convert(new VMLShapeTypeMapping(this._ctx, this._writer));

            foreach (var VMLEntry in VMLEntriesList)
            {                

                bse = (BlipStoreEntry)VMLEntry[0];
                var options = (ShapeOptions)VMLEntry[2];
                var bounds = (Rectangle)VMLEntry[3];
                string spid = (string)VMLEntry[4];
                var size = (Point)VMLEntry[5];

                var imgPart = copyPicture(bse, ref size);
                if (imgPart != null)
                {

                    //v:shape
                    this._writer.WriteStartElement("v", "shape", OpenXmlNamespaces.VectorML);
                    this._writer.WriteAttributeString("id", spid);
                    this._writer.WriteAttributeString("type", "#" + VMLShapeTypeMapping.GenerateTypeId(type));

                    var style = new StringBuilder();


                    style.Append("position:absolute;");
                    style.Append("left:" + (new EmuValue(Utils.MasterCoordToEMU(bounds.Left)).ToPoints()).ToString() + "pt;");
                    style.Append("top:" + (new EmuValue(Utils.MasterCoordToEMU(bounds.Top)).ToPoints()).ToString() + "pt;");
                    style.Append("width:").Append(new EmuValue(Utils.MasterCoordToEMU(bounds.Width)).ToPoints()).Append("pt;");
                    style.Append("height:").Append(new EmuValue(Utils.MasterCoordToEMU(bounds.Height)).ToPoints()).Append("pt;");
                    this._writer.WriteAttributeString("style", style.ToString());

                    foreach (var entry in options.OptionsByID.Values)
                    {
                        switch (entry.pid)
                        {
                            //BORDERS

                            case ShapeOptions.PropertyId.borderBottomColor:
                                var bottomColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                                this._writer.WriteAttributeString("o", "borderbottomcolor", OpenXmlNamespaces.Office, "#" + bottomColor.SixDigitHexCode);
                                break;
                            case ShapeOptions.PropertyId.borderLeftColor:
                                var leftColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                                this._writer.WriteAttributeString("o", "borderleftcolor", OpenXmlNamespaces.Office, "#" + leftColor.SixDigitHexCode);
                                break;
                            case ShapeOptions.PropertyId.borderRightColor:
                                var rightColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                                this._writer.WriteAttributeString("o", "borderrightcolor", OpenXmlNamespaces.Office, "#" + rightColor.SixDigitHexCode);
                                break;
                            case ShapeOptions.PropertyId.borderTopColor:
                                var topColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                                this._writer.WriteAttributeString("o", "bordertopcolor", OpenXmlNamespaces.Office, "#" + topColor.SixDigitHexCode);
                                break;
                        }
                    }

                    //v:imageData
                    this._writer.WriteStartElement("v", "imagedata", OpenXmlNamespaces.VectorML);
                    this._writer.WriteAttributeString("o", "relid", OpenXmlNamespaces.Office, imgPart.RelIdToString);
                    this._writer.WriteAttributeString("o", "title", OpenXmlNamespaces.Office, "");
                    this._writer.WriteEndElement(); //imagedata

                    //close v:shape
                    this._writer.WriteEndElement();                   
                }
            }

            this._writer.WriteEndElement(); //xml
            this._writer.WriteEndDocument();
            this._writer.Flush();
        }

        /// <summary>
        /// Copies the picture from the binary stream to the zip archive 
        /// and creates the relationships for the image.
        /// </summary>
        /// <param name="pict">The PictureDescriptor</param>
        /// <returns>The created ImagePart</returns>
        protected ImagePart copyPicture(BlipStoreEntry bse, ref Point size)
        {
            //create the image part
            ImagePart imgPart = null;
            if (bse != null)
            {
                switch (bse.btWin32)
                {
                    case BlipStoreEntry.BlipType.msoblipEMF:
                        imgPart = this._targetPart.AddImagePart(ImagePart.ImageType.Emf);
                        break;
                    case BlipStoreEntry.BlipType.msoblipWMF:
                        imgPart = this._targetPart.AddImagePart(ImagePart.ImageType.Wmf);
                        break;
                    case BlipStoreEntry.BlipType.msoblipJPEG:
                    case BlipStoreEntry.BlipType.msoblipCMYKJPEG:
                        imgPart = this._targetPart.AddImagePart(ImagePart.ImageType.Jpeg);
                        break;
                    case BlipStoreEntry.BlipType.msoblipPNG:
                        imgPart = this._targetPart.AddImagePart(ImagePart.ImageType.Png);
                        break;
                    case BlipStoreEntry.BlipType.msoblipTIFF:
                        imgPart = this._targetPart.AddImagePart(ImagePart.ImageType.Tiff);
                        break;
                    case BlipStoreEntry.BlipType.msoblipDIB:
                    case BlipStoreEntry.BlipType.msoblipPICT:
                    case BlipStoreEntry.BlipType.msoblipERROR:
                    case BlipStoreEntry.BlipType.msoblipUNKNOWN:
                    case BlipStoreEntry.BlipType.msoblipLastClient:
                    case BlipStoreEntry.BlipType.msoblipFirstClient:
                        //throw new Exception("Cannot convert picture of type " + bse.btWin32);
                        break;
                }

                imgPart.TargetDirectory = "..\\media";
                var outStream = imgPart.GetStream();

                var mb = this._ctx.Ppt.PicturesContainer._pictures[bse.foDelay];

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
                            outStream.Write(decompressed, 0, decompressed.Length);

                            break;
                        case BlipStoreEntry.BlipType.msoblipJPEG:
                        case BlipStoreEntry.BlipType.msoblipCMYKJPEG:
                        case BlipStoreEntry.BlipType.msoblipPNG:
                        case BlipStoreEntry.BlipType.msoblipTIFF:

                            //it's a bitmap image
                            var bitBlip = (BitmapBlip)mb;
                            outStream.Write(bitBlip.m_pvBits, 0, bitBlip.m_pvBits.Length);

                            break;
                    }
                }
            }
            return imgPart;
        }
    }
}
