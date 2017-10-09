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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Writer;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class MacroBinaryMapping : AbstractOpenXmlMapping,
        IMapping<XlsDocument>
    {
        private ExcelContext ctx;
        private string projectFolder = "\\_VBA_PROJECT_CUR";
        private string vbaFolder = "\\_VBA_PROJECT_CUR\\VBA";
        private string projectFile = "\\_VBA_PROJECT_CUR\\PROJECT";
        private string projectWmFile = "\\_VBA_PROJECT_CUR\\PROJECTwm";

        public MacroBinaryMapping(ExcelContext ctx)
            : base(null)
        {
            this.ctx = ctx;
        }

        public void Apply(XlsDocument xls)
        {
            //get the Class IDs of the directories
            Guid macroClsid = new Guid();
            Guid vbaClsid = new Guid();
            foreach (DirectoryEntry entry in xls.Storage.AllEntries)
            {
                if (entry.Path == projectFolder)
                {
                    macroClsid = entry.ClsId;
                }
                else if (entry.Path == vbaFolder)
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
            foreach (DirectoryEntry entry in xls.Storage.AllStreamEntries)
            {
                if (entry.Path.StartsWith(vbaFolder))
                {
                    vba.AddStreamDirectoryEntry(entry.Name, xls.Storage.GetStream(entry.Path));
                }
            }

            //copy the project streams
            storage.RootDirectoryEntry.AddStreamDirectoryEntry("PROJECT", xls.Storage.GetStream(projectFile));
            storage.RootDirectoryEntry.AddStreamDirectoryEntry("PROJECTwm", xls.Storage.GetStream(projectWmFile));

           //write the storage to the xml part
            storage.write(ctx.SpreadDoc.WorkbookPart.VbaProjectPart.GetStream());
        }
    }
}
