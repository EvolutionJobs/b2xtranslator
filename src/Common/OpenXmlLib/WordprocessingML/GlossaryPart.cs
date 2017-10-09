using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class GlossaryPart : MainDocumentPart
    {
        public GlossaryPart(OpenXmlPartContainer parent, string contentType)
            : base (parent, contentType)
        {
            
        }

        public override string RelationshipType { get { return OpenXmlRelationshipTypes.GlossaryDocument; } }
        public override string TargetDirectory { get { return "glossary"; } }
    }
}
