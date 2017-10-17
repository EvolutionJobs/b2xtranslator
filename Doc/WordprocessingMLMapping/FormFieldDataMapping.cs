using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class FormFieldDataMapping :
        AbstractOpenXmlMapping,
        IMapping<FormFieldData>
    {
        private const int UNDEFINED_RESULT = 25;

        public FormFieldDataMapping(XmlWriter writer)
            : base(writer)
        {
        }

        public void Apply(FormFieldData ffd)
        {
            this._writer.WriteStartElement("w", "ffData", OpenXmlNamespaces.WordprocessingML);

            //name
            this._writer.WriteStartElement("w", "name", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzName);
            this._writer.WriteEndElement();

            //calcOnExit
            this._writer.WriteStartElement("w", "calcOnExit", OpenXmlNamespaces.WordprocessingML);
            if (ffd.fRecalc)
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "1");
            else
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, "0");
            this._writer.WriteEndElement();

            //entry macro
            if (ffd.xstzEntryMcr != null && ffd.xstzEntryMcr.Length > 0)
            {
                this._writer.WriteStartElement("w", "entryMacro", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzEntryMcr);
                this._writer.WriteEndElement();
            }

            //exit macro
            if (ffd.xstzExitMcr != null && ffd.xstzExitMcr.Length > 0)
            {
                this._writer.WriteStartElement("w", "exitMacro", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzExitMcr);
                this._writer.WriteEndElement();
            }

            //help text
            if (ffd.xstzHelpText != null && ffd.xstzHelpText.Length > 0)
            {
                this._writer.WriteStartElement("w", "helpText", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "text");
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzHelpText);
                this._writer.WriteEndElement();
            }

            //status text
            if (ffd.xstzStatText != null && ffd.xstzStatText.Length > 0)
            {
                this._writer.WriteStartElement("w", "statusText", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "type", OpenXmlNamespaces.WordprocessingML, "text");
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzStatText);
                this._writer.WriteEndElement();
            }

            //start custom properties
            switch (ffd.iType)
            {
                case FormFieldData.FormFieldType.iTypeText:
                    this._writer.WriteStartElement("w", "textInput", OpenXmlNamespaces.WordprocessingML);

                    //type
                    if (ffd.iTypeTxt != FormFieldData.TextboxType.regular)
                    {
                        this._writer.WriteStartElement("w", "type", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.iTypeTxt.ToString());
                        this._writer.WriteEndElement();
                    }

                    //length
                    if (ffd.cch > 0)
                    {
                        this._writer.WriteStartElement("w", "maxLength", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.cch.ToString());
                        this._writer.WriteEndElement();
                    }
                    
                    //textformat
                    if (ffd.xstzTextFormat != null && ffd.xstzTextFormat.Length > 0)
                    {
                        this._writer.WriteStartElement("w", "format", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzTextFormat);
                        this._writer.WriteEndElement();
                    }

                    //default text
                    if (ffd.xstzTextDef != null && ffd.xstzTextDef.Length > 0)
                    {
                        this._writer.WriteStartElement("w", "default", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.xstzTextDef);
                        this._writer.WriteEndElement();
                    }

                    break;
                case FormFieldData.FormFieldType.iTypeChck:
                    this._writer.WriteStartElement("w", "checkBox", OpenXmlNamespaces.WordprocessingML);

                    //checked <w:checked w:val="0"/>
                    if (ffd.iRes != UNDEFINED_RESULT)
                    {
                        this._writer.WriteStartElement("w", "checked", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.iRes.ToString());
                        this._writer.WriteEndElement();
                    }

                    //size 
                    if (ffd.hps >= 2 && ffd.hps <= 3168)
                    {
                        this._writer.WriteStartElement("w", "size", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.hps.ToString());
                        this._writer.WriteEndElement();
                    }
                    else
                    {
                        this._writer.WriteStartElement("w", "sizeAuto", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteEndElement();
                    }

                    //default setting
                    this._writer.WriteStartElement("w", "default", OpenXmlNamespaces.WordprocessingML);
                    this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.wDef.ToString());
                    this._writer.WriteEndElement();
                    
                    break;
                case FormFieldData.FormFieldType.iTypeDrop:
                    this._writer.WriteStartElement("w", "ddList", OpenXmlNamespaces.WordprocessingML);

                    //selected item
                    if (ffd.iRes != UNDEFINED_RESULT)
                    {
                        this._writer.WriteStartElement("w", "result", OpenXmlNamespaces.WordprocessingML);
                        this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, ffd.iRes.ToString());
                        this._writer.WriteEndElement();
                    }

                    //default

                    //_entries

                    break;
                default:
                    this._writer.WriteStartElement("w", "textInput", OpenXmlNamespaces.WordprocessingML);

                    break;
            }

            //close custom properties
            this._writer.WriteEndElement();

            //close ffData
            this._writer.WriteEndElement();
        }
    }
}
