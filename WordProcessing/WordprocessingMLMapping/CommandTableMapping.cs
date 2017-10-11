using System;
using b2xtranslator.DocFileFormat;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib;
using System.Xml;
using System.IO;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class CommandTableMapping : AbstractOpenXmlMapping,
        IMapping<CommandTable>
    {
        private CommandTable _tcg;
        private ConversionContext _ctx;

        public CommandTableMapping(ConversionContext ctx)
            : base(XmlWriter.Create(ctx.Docx.MainDocumentPart.CustomizationsPart.GetStream(), ctx.WriterSettings))
        {
            this._ctx = ctx;
        }

        public void Apply(CommandTable tcg)
        {
            this._tcg = tcg;
            this._writer.WriteStartElement("wne", "tcg", OpenXmlNamespaces.MicrosoftWordML);

            //write the keymaps
            this._writer.WriteStartElement("wne", "keymaps", OpenXmlNamespaces.MicrosoftWordML);
            for (int i = 0; i < tcg.KeyMapEntries.Count; i++)
            {
                writeKeyMapEntry(tcg.KeyMapEntries[i]);
            }
            this._writer.WriteEndElement();

            //write the toolbars
            if (tcg.CustomToolbars != null)
            {
                this._writer.WriteStartElement("wne", "toolbars", OpenXmlNamespaces.MicrosoftWordML);
                writeToolbar(tcg.CustomToolbars);
                this._writer.WriteEndElement();
            }

            this._writer.WriteEndElement();

            this._writer.Flush();
        }


        private void writeToolbar(CustomToolbarWrapper toolbars)
        {
            //write the xml
            this._writer.WriteStartElement("wne", "toolbarData", OpenXmlNamespaces.MicrosoftWordML);
            this._writer.WriteAttributeString("r", "id", 
                OpenXmlNamespaces.Relationships,
                this._ctx.Docx.MainDocumentPart.CustomizationsPart.ToolbarsPart.RelIdToString
             );
            this._writer.WriteEndElement();

            //copy the toolbar
            var s = this._ctx.Docx.MainDocumentPart.CustomizationsPart.ToolbarsPart.GetStream();
            s.Write(toolbars.RawBytes, 0, toolbars.RawBytes.Length);
        }


        private void writeKeyMapEntry(KeyMapEntry kme)
        {
            this._writer.WriteStartElement("wne", "keymap", OpenXmlNamespaces.MicrosoftWordML);

            //primary KCM
            if (kme.kcm1 > 0)
            {
                this._writer.WriteAttributeString("wne", "kcmPrimary",
                    OpenXmlNamespaces.MicrosoftWordML,
                    string.Format("{0:x4}", kme.kcm1));
            }

            this._writer.WriteStartElement("wne", "macro", OpenXmlNamespaces.MicrosoftWordML);

            this._writer.WriteAttributeString("wne", "macroName",
                OpenXmlNamespaces.MicrosoftWordML,
                this._tcg.MacroNames[kme.paramCid.ibstMacro]
                );

            this._writer.WriteEndElement();

            this._writer.WriteEndElement();
        }
    }
}
