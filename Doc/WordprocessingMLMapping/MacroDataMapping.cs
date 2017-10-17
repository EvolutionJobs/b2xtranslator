using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class MacroDataMapping : DocumentMapping
    {
        public MacroDataMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart.VbaProjectPart.VbaDataPart)
        {
            this._ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            var reader = new VirtualStreamReader(doc.Storage.GetStream("\\Macros\\PROJECTwm"));
            this._writer.WriteStartElement("wne", "vbaSuppData", OpenXmlNamespaces.MicrosoftWordML);

            this._writer.WriteStartElement("wne", "mcds", OpenXmlNamespaces.MicrosoftWordML);
            for (int i = 0; i < doc.CommandTable.MacroDatas.Count; i++)
            {
                this._writer.WriteStartElement("wne", "mcd", OpenXmlNamespaces.MicrosoftWordML);
                var mcd = doc.CommandTable.MacroDatas[i];

                if (doc.CommandTable.MacroNames != null)
                {
                    this._writer.WriteAttributeString(
                        "wne", "macroName",
                        OpenXmlNamespaces.MicrosoftWordML,
                        doc.CommandTable.MacroNames[mcd.ibst]);
                }

                if (doc.CommandTable.CommandStringTable != null)
                {
                    this._writer.WriteAttributeString(
                        "wne", "name",
                        OpenXmlNamespaces.MicrosoftWordML,
                        doc.CommandTable.CommandStringTable.Strings[mcd.ibstName]);
                }

                this._writer.WriteEndElement();
            }
            this._writer.WriteEndElement();

            this._writer.WriteEndElement();
            reader.Close();

            this._writer.Flush();
        }
    }
}
