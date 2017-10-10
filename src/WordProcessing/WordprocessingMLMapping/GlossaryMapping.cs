using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class GlossaryMapping : DocumentMapping
    {
        public GlossaryMapping(ConversionContext ctx, ContentPart targetPart)
            : base(ctx, targetPart)
        {

        }

        public override void Apply(WordDocument doc)
        {
            _doc = doc;

            //start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("w", "glossaryDocument", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "docParts", OpenXmlNamespaces.WordprocessingML);

            for (int i = 0; i < _doc.AutoTextPlex.CharacterPositions.Count - 2; i++)
            {
                int cpStart = _doc.AutoTextPlex.CharacterPositions[i];
                int cpEnd = _doc.AutoTextPlex.CharacterPositions[i+1];

                writeAutoTextDocPart(cpStart, cpEnd, i);
            }

            //end the document
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndDocument();

            _writer.Flush();
        }

        private void writeAutoTextDocPart(int startCp, int endCp, int index)
        {
            _writer.WriteStartElement("w", "docPart", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "docPartPr", OpenXmlNamespaces.WordprocessingML);

            //write the name
            _writer.WriteStartElement("w", "name", OpenXmlNamespaces.WordprocessingML);
            string name = _doc.AutoTextNames.Strings[index];
            if((int)name[name.Length-1] == 1)
            {
                name = name.Remove(name.Length-1);
            }
            _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, name);
            _writer.WriteEndElement();

            //write the category
            _writer.WriteStartElement("w", "category", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "name", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "General");
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "gallery", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "autoTxt");
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            //write behaviors
            _writer.WriteStartElement("w", "behaviors", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteStartElement("w", "behavior", OpenXmlNamespaces.WordprocessingML);
            _writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "content");
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "docPartBody", OpenXmlNamespaces.WordprocessingML);

            writeParagraph(startCp, endCp, false);

            _writer.WriteEndElement();
            _writer.WriteEndElement();
        }
    }
}
