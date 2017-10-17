using System;
using b2xtranslator.DocFileFormat;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.StructuredStorage.Writer;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class MacroBinaryMapping : DocumentMapping
    {
        public MacroBinaryMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart.VbaProjectPart)
        {
            this._ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            //get the Class IDs of the directories
            var macroClsid = new Guid();
            var vbaClsid = new Guid();
            foreach (var entry in doc.Storage.AllEntries)
            {
                if (entry.Path == "\\Macros")
                {
                    macroClsid = entry.ClsId;
                }
                else if(entry.Path == "\\Macros\\VBA")
                {
                    vbaClsid = entry.ClsId;
                }
            }

            //create a new storage
            var storage = new StructuredStorageWriter();
            storage.RootDirectoryEntry.setClsId(macroClsid);

            //copy the VBA directory
            var vba = storage.RootDirectoryEntry.AddStorageDirectoryEntry("VBA");
            vba.setClsId(vbaClsid);
            foreach (var entry in doc.Storage.AllStreamEntries)
            {
                if (entry.Path.StartsWith("\\Macros\\VBA"))
                {
                    vba.AddStreamDirectoryEntry(entry.Name, doc.Storage.GetStream(entry.Path));
                }
            }

            //copy the project streams
            storage.RootDirectoryEntry.AddStreamDirectoryEntry("PROJECT", doc.Storage.GetStream("\\Macros\\PROJECT"));
            storage.RootDirectoryEntry.AddStreamDirectoryEntry("PROJECTwm", doc.Storage.GetStream("\\Macros\\PROJECTwm"));

           //write the storage to the xml part
            storage.write(this._targetPart.GetStream());
        }
    }
}
