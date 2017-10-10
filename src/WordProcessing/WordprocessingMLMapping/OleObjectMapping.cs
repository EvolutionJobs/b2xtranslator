using System;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.DocFileFormat;
using b2xtranslator.StructuredStorage.Writer;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class OleObjectMapping :
        AbstractOpenXmlMapping,
        IMapping<OleObject>
    {
        ContentPart _targetPart;
        WordDocument _doc;
        PictureDescriptor _pict;

        public OleObjectMapping(XmlWriter writer, WordDocument doc, ContentPart targetPart, PictureDescriptor pict)
            : base(writer)
        {
            this._targetPart = targetPart;
            this._doc = doc;
            this._pict = pict;
        }

        public void Apply(OleObject ole)
        {
            this._writer.WriteStartElement("o", "OLEObject", OpenXmlNamespaces.Office);

            EmbeddedObjectPart.ObjectType type;
            if (ole.ClipboardFormat == "Biff8")
            {
                type = EmbeddedObjectPart.ObjectType.Excel;
            }
            else if (ole.ClipboardFormat == "MSWordDoc")
            {
                type = EmbeddedObjectPart.ObjectType.Word;
            }
            else if (ole.ClipboardFormat == "MSPresentation")
            {
                type = EmbeddedObjectPart.ObjectType.Powerpoint;
            }
            else
            {
                type = EmbeddedObjectPart.ObjectType.Other;
            }

            //type
            if (ole.fLinked)
            {
                var link = new Uri(ole.Link);
                var rel = this._targetPart.AddExternalRelationship(OpenXmlRelationshipTypes.OleObject, link);
                this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, rel.Id);
                this._writer.WriteAttributeString("Type", "Link");
                this._writer.WriteAttributeString("UpdateMode", ole.UpdateMode.ToString());
            }
            else
            {
                var part = this._targetPart.AddEmbeddedObjectPart(type);
                this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, part.RelIdToString);
                this._writer.WriteAttributeString("Type", "Embed");

                //copy the object
                copyEmbeddedObject(ole, part);
            }

            //ProgID
            this._writer.WriteAttributeString("ProgID", ole.Program);

            //ShapeId
            this._writer.WriteAttributeString("ShapeID", this._pict.ShapeContainer.GetHashCode().ToString());

            //DrawAspect
            this._writer.WriteAttributeString("DrawAspect", "Content");

            //ObjectID
            this._writer.WriteAttributeString("ObjectID", ole.ObjectId);

            this._writer.WriteEndElement();
        }


        /// <summary>
        /// Writes the embedded OLE object from the ObjectPool of the binary file to the OpenXml Package.
        /// </summary>
        /// <param name="ole"></param>
        private void copyEmbeddedObject(OleObject ole, EmbeddedObjectPart part)
        {
            //create a new storage
            var writer = new StructuredStorageWriter();

            // Word will not open embedded charts if a CLSID is set.
            if(ole.Program.StartsWith("Excel.Chart") == false)
            {
                writer.RootDirectoryEntry.setClsId(ole.ClassId);
            }

            //copy the OLE streams from the old storage to the new storage
            foreach (string oleStream in ole.Streams.Keys)
            {
                writer.RootDirectoryEntry.AddStreamDirectoryEntry(oleStream, ole.Streams[oleStream]);
            }

           //write the storage to the xml part
           writer.write(part.GetStream());
        }
    }
}
