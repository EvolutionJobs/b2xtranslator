using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class CommentsMapping : DocumentMapping
    {
        public CommentsMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart.CommentsPart)
        {
            this._ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            this._doc = doc;
            int index = 0;

            this._writer.WriteStartElement("w", "comments", OpenXmlNamespaces.WordprocessingML);

            int cp = doc.FIB.ccpText + doc.FIB.ccpFtn + doc.FIB.ccpHdr;
            for (int i = 0; i < doc.AnnotationsReferencePlex.Elements.Count; i++)
			{
                this._writer.WriteStartElement("w", "comment", OpenXmlNamespaces.WordprocessingML);

                var atrdPre10 = (AnnotationReferenceDescriptor)doc.AnnotationsReferencePlex.Elements[index];
                this._writer.WriteAttributeString("w", "id", OpenXmlNamespaces.WordprocessingML, index.ToString());
                this._writer.WriteAttributeString("w", "author", OpenXmlNamespaces.WordprocessingML, doc.AnnotationOwners[atrdPre10.AuthorIndex]);
                this._writer.WriteAttributeString("w", "initials", OpenXmlNamespaces.WordprocessingML, atrdPre10.UserInitials);

                //ATRDpost10 is optional and not saved in all files
                if (doc.AnnotationReferenceExtraTable != null && 
                    doc.AnnotationReferenceExtraTable.Count > index)
                {
                    var atrdPost10 = doc.AnnotationReferenceExtraTable[index];
                    atrdPost10.Date.Convert(new DateMapping(this._writer));
                }

                cp = writeParagraph(cp);
                this._writer.WriteEndElement();
                index++;
            }

            this._writer.WriteEndElement();

            this._writer.Flush();
        }
    }
}
