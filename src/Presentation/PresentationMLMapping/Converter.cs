using System;
using System.Text;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.PptFileFormat;
using System.IO;
using b2xtranslator.OpenXmlLib.PresentationML;
using System.Xml;

namespace b2xtranslator.PresentationMLMapping
{
    public class Converter
    {
        public static OpenXmlPackage.DocumentType DetectOutputType(PowerpointDocument ppt)
        {
            var returnType = OpenXmlPackage.DocumentType.Document;

            try
            {
                //ToDo: Find better way to detect macro type
                if (ppt.VbaProject != null)
                {
                    returnType = OpenXmlPackage.DocumentType.MacroEnabledDocument;
                }
            }
            catch (Exception)
            {
            }

            return returnType;
        }

        public static string GetConformFilename(string choosenFilename, OpenXmlPackage.DocumentType outType)
        {
            string outExt = ".pptx";
            switch (outType)
            {
                case OpenXmlPackage.DocumentType.Document:
                    outExt = ".pptx";
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledDocument:
                    outExt = ".pptm";
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledTemplate:
                    outExt = ".potm";
                    break;
                case OpenXmlPackage.DocumentType.Template:
                    outExt = ".potx";
                    break;
                default:
                    outExt = ".pptx";
                    break;
            }

            string inExt = Path.GetExtension(choosenFilename);
            if (inExt != null)
            {
                return choosenFilename.Replace(inExt, outExt);
            }
            else
            {
                return choosenFilename + outExt;
            }
        }

        public static void Convert(PowerpointDocument ppt, PresentationDocument pptx)
        {
            using (pptx)
            {
                // Setup the writer
                var xws = new XmlWriterSettings();
                xws.OmitXmlDeclaration = false;
                xws.CloseOutput = true;
                xws.Encoding = Encoding.UTF8;
                xws.ConformanceLevel = ConformanceLevel.Document;

                // Setup the context
                var context = new ConversionContext(ppt);
                context.WriterSettings = xws;
                context.Pptx = pptx;

                // Write presentation.xml
                ppt.Convert(new PresentationPartMapping(context));

                //AppMapping app = new AppMapping(pptx.AddAppPropertiesPart(), xws);
                //app.Apply(null);

                //CoreMapping core = new CoreMapping(pptx.AddCoreFilePropertiesPart(), xws);
                //core.Apply(null);

            }
        }
    }
}
