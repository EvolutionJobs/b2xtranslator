

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib;
using System.Xml;

namespace b2xtranslator.PresentationMLMapping
{
    public abstract class PresentationMapping<T> :
        AbstractOpenXmlMapping,
        IMapping<T>
        where T : IVisitable
    {
        protected ConversionContext _ctx;
        public ContentPart targetPart;
        
        public PresentationMapping(ConversionContext ctx, ContentPart targetPart)
            : base(XmlWriter.Create(targetPart.GetStream(), ctx.WriterSettings))
        {
            this._ctx = ctx;
            this.targetPart = targetPart;
        }

        public abstract void Apply(T mapElement);
    }
}
