using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class VbaProjectMapping : AbstractOpenXmlMapping,
        IMapping<ExOleObjStgAtom>
    {
        private VbaProjectPart _targetPart;

        public VbaProjectMapping(VbaProjectPart targetPart)
            : base(null)
        {
            _targetPart = targetPart;
        }

        public void Apply(ExOleObjStgAtom vbaProject)
        {
            byte[] bytes = vbaProject.DecompressData();
            _targetPart.GetStream().Write(bytes, 0, bytes.Length);
            
        }
    }
}
