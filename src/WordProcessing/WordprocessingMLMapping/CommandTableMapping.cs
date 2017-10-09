using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class CommandTableMapping : AbstractOpenXmlMapping,
        IMapping<CommandTable>
    {
        private CommandTable _tcg;
        private ConversionContext _ctx;

        public CommandTableMapping(ConversionContext ctx)
            : base(XmlWriter.Create(ctx.Docx.MainDocumentPart.CustomizationsPart.GetStream(), ctx.WriterSettings))
        {
            _ctx = ctx;
        }

        public void Apply(CommandTable tcg)
        {
            _tcg = tcg;
            _writer.WriteStartElement("wne", "tcg", OpenXmlNamespaces.MicrosoftWordML);

            //write the keymaps
            _writer.WriteStartElement("wne", "keymaps", OpenXmlNamespaces.MicrosoftWordML);
            for (int i = 0; i < tcg.KeyMapEntries.Count; i++)
            {
                writeKeyMapEntry(tcg.KeyMapEntries[i]);
            }
            _writer.WriteEndElement();

            //write the toolbars
            if (tcg.CustomToolbars != null)
            {
                _writer.WriteStartElement("wne", "toolbars", OpenXmlNamespaces.MicrosoftWordML);
                writeToolbar(tcg.CustomToolbars);
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();

            _writer.Flush();
        }


        private void writeToolbar(CustomToolbarWrapper toolbars)
        {
            //write the xml
            _writer.WriteStartElement("wne", "toolbarData", OpenXmlNamespaces.MicrosoftWordML);
            _writer.WriteAttributeString("r", "id", 
                OpenXmlNamespaces.Relationships,
                _ctx.Docx.MainDocumentPart.CustomizationsPart.ToolbarsPart.RelIdToString
             );
            _writer.WriteEndElement();

            //copy the toolbar
            Stream s = _ctx.Docx.MainDocumentPart.CustomizationsPart.ToolbarsPart.GetStream();
            s.Write(toolbars.RawBytes, 0, toolbars.RawBytes.Length);
        }


        private void writeKeyMapEntry(KeyMapEntry kme)
        {
            _writer.WriteStartElement("wne", "keymap", OpenXmlNamespaces.MicrosoftWordML);

            //primary KCM
            if (kme.kcm1 > 0)
            {
                _writer.WriteAttributeString("wne", "kcmPrimary",
                    OpenXmlNamespaces.MicrosoftWordML,
                    String.Format("{0:x4}", kme.kcm1));
            }

            _writer.WriteStartElement("wne", "macro", OpenXmlNamespaces.MicrosoftWordML);

            _writer.WriteAttributeString("wne", "macroName",
                OpenXmlNamespaces.MicrosoftWordML,
                _tcg.MacroNames[kme.paramCid.ibstMacro]
                );

            _writer.WriteEndElement();

            _writer.WriteEndElement();
        }
    }
}
