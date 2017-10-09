/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.IO;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Writer;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
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
            _targetPart = targetPart;
            _doc = doc;
            _pict = pict;
        }

        public void Apply(OleObject ole)
        {
            _writer.WriteStartElement("o", "OLEObject", OpenXmlNamespaces.Office);

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
                Uri link = new Uri(ole.Link);
                ExternalRelationship rel = _targetPart.AddExternalRelationship(OpenXmlRelationshipTypes.OleObject, link);
                _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, rel.Id);
                _writer.WriteAttributeString("Type", "Link");
                _writer.WriteAttributeString("UpdateMode", ole.UpdateMode.ToString());
            }
            else
            {
                EmbeddedObjectPart part = _targetPart.AddEmbeddedObjectPart(type);
                _writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, part.RelIdToString);
                _writer.WriteAttributeString("Type", "Embed");

                //copy the object
                copyEmbeddedObject(ole, part);
            }

            //ProgID
            _writer.WriteAttributeString("ProgID", ole.Program);

            //ShapeId
            _writer.WriteAttributeString("ShapeID", _pict.ShapeContainer.GetHashCode().ToString());

            //DrawAspect
            _writer.WriteAttributeString("DrawAspect", "Content");

            //ObjectID
            _writer.WriteAttributeString("ObjectID", ole.ObjectId);

            _writer.WriteEndElement();
        }


        /// <summary>
        /// Writes the embedded OLE object from the ObjectPool of the binary file to the OpenXml Package.
        /// </summary>
        /// <param name="ole"></param>
        private void copyEmbeddedObject(OleObject ole, EmbeddedObjectPart part)
        {
            //create a new storage
            StructuredStorageWriter writer = new StructuredStorageWriter();

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
