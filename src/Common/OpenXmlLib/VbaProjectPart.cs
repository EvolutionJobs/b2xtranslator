using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib
{
    public class VbaProjectPart : ContentPart
    {
        internal VbaProjectPart(OpenXmlPartContainer parent)
            : base(parent, 0)
        {
        }

        public override string ContentType
        {
            get { return MicrosoftWordContentTypes.VbaProject; }
        }

        public override string RelationshipType
        {
            get { return MicrosoftWordRelationshipTypes.VbaProject; }
        }

        public override string TargetName { get { return "vbaProject"; } }
        public override string TargetExt { get { return ".bin"; } }
        public override string TargetDirectory { get { return ""; } }

        protected VbaDataPart _vbaDataPart;
        public VbaDataPart VbaDataPart
        {
            get
            {
                if(_vbaDataPart == null)
                {
                    _vbaDataPart = this.AddPart(new VbaDataPart(this));
                }
                return _vbaDataPart;
            }
            
        }
    }
}
