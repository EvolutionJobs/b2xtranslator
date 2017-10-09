/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Writer;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class MacroBinaryMapping : DocumentMapping
    {
        public MacroBinaryMapping(ConversionContext ctx)
            : base(ctx, ctx.Docx.MainDocumentPart.VbaProjectPart)
        {
            _ctx = ctx;
        }

        public override void Apply(WordDocument doc)
        {
            //get the Class IDs of the directories
            Guid macroClsid = new Guid();
            Guid vbaClsid = new Guid();
            foreach (DirectoryEntry entry in doc.Storage.AllEntries)
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
            StructuredStorageWriter storage = new StructuredStorageWriter();
            storage.RootDirectoryEntry.setClsId(macroClsid);

            //copy the VBA directory
            StorageDirectoryEntry vba = storage.RootDirectoryEntry.AddStorageDirectoryEntry("VBA");
            vba.setClsId(vbaClsid);
            foreach (DirectoryEntry entry in doc.Storage.AllStreamEntries)
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
            storage.write(_targetPart.GetStream());
        }
    }
}
