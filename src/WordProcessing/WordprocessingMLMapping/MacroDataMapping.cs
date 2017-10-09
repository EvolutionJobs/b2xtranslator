using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class MacroDataMapping : DocumentMapping
    {
        public MacroDataMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart.VbaProjectPart.VbaDataPart)
        {
            _ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            VirtualStreamReader reader = new VirtualStreamReader(doc.Storage.GetStream("\\Macros\\PROJECTwm"));
            _writer.WriteStartElement("wne", "vbaSuppData", OpenXmlNamespaces.MicrosoftWordML);

            _writer.WriteStartElement("wne", "mcds", OpenXmlNamespaces.MicrosoftWordML);
            for (int i = 0; i < doc.CommandTable.MacroDatas.Count; i++)
            {
                _writer.WriteStartElement("wne", "mcd", OpenXmlNamespaces.MicrosoftWordML);
                MacroData mcd = doc.CommandTable.MacroDatas[i];

                if (doc.CommandTable.MacroNames != null)
                {
                    _writer.WriteAttributeString(
                        "wne", "macroName",
                        OpenXmlNamespaces.MicrosoftWordML,
                        doc.CommandTable.MacroNames[mcd.ibst]);
                }

                if (doc.CommandTable.CommandStringTable != null)
                {
                    _writer.WriteAttributeString(
                        "wne", "name",
                        OpenXmlNamespaces.MicrosoftWordML,
                        doc.CommandTable.CommandStringTable.Strings[mcd.ibstName]);
                }

                _writer.WriteEndElement();
            }
            _writer.WriteEndElement();

            _writer.WriteEndElement();
            reader.Close();

            _writer.Flush();
        }
    }
}
