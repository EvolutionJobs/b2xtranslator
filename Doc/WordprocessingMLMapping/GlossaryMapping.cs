using b2xtranslator.OpenXmlLib;
using b2xtranslator.DocFileFormat;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class GlossaryMapping : DocumentMapping
    {
        public GlossaryMapping(ConversionContext ctx, ContentPart targetPart)
            : base(ctx, targetPart)
        {

        }

        public override void Apply(WordDocument doc)
        {
            this._doc = doc;

            //start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("w", "glossaryDocument", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteStartElement("w", "docParts", OpenXmlNamespaces.WordprocessingML);

            for (int i = 0; i < this._doc.AutoTextPlex.CharacterPositions.Count - 2; i++)
            {
                int cpStart = this._doc.AutoTextPlex.CharacterPositions[i];
                int cpEnd = this._doc.AutoTextPlex.CharacterPositions[i+1];

                writeAutoTextDocPart(cpStart, cpEnd, i);
            }

            //end the document
            this._writer.WriteEndElement();
            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }

        private void writeAutoTextDocPart(int startCp, int endCp, int index)
        {
            this._writer.WriteStartElement("w", "docPart", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteStartElement("w", "docPartPr", OpenXmlNamespaces.WordprocessingML);

            //write the name
            this._writer.WriteStartElement("w", "name", OpenXmlNamespaces.WordprocessingML);
            string name = this._doc.AutoTextNames.Strings[index];
            if((int)name[name.Length-1] == 1)
            {
                name = name.Remove(name.Length-1);
            }
            this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, name);
            this._writer.WriteEndElement();

            //write the category
            this._writer.WriteStartElement("w", "category", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteStartElement("w", "name", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "General");
            this._writer.WriteEndElement();
            this._writer.WriteStartElement("w", "gallery", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "autoTxt");
            this._writer.WriteEndElement();
            this._writer.WriteEndElement();

            //write behaviors
            this._writer.WriteStartElement("w", "behaviors", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteStartElement("w", "behavior", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "content");
            this._writer.WriteEndElement();
            this._writer.WriteEndElement();

            this._writer.WriteEndElement();
            this._writer.WriteStartElement("w", "docPartBody", OpenXmlNamespaces.WordprocessingML);

            writeParagraph(startCp, endCp, false);

            this._writer.WriteEndElement();
            this._writer.WriteEndElement();
        }
    }
}
