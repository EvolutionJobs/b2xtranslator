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
using System.Globalization;
using System.IO;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Shell;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.ZipUtils;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.SpreadsheetML;

namespace DIaLOGIKa.b2xtranslator.xls2x
{
    class Program : CommandLineTranslator
    {
        public static string ToolName = "xls2x";
        public static string RevisionResource = "DIaLOGIKa.b2xtranslator.xls2x.revision.txt";
        public static string ContextMenuInputExtension = ".xls";
        public static string ContextMenuText = "Convert to .xlsx";

        static void Main(string[] args)
        {
            ParseArgs(args, ToolName);

            InitializeLogger();

            PrintWelcome(ToolName, RevisionResource);

            try
            {
                //copy processing file
                var procFile = new ProcessingFile(InputFile);

                //make output file name
                if (ChoosenOutputFile == null)
                {
                    if (InputFile.Contains("."))
                    {
                        ChoosenOutputFile = InputFile.Remove(InputFile.LastIndexOf(".")) + ".xlsx";
                    }
                    else
                    {
                        ChoosenOutputFile = InputFile + ".xlsx";
                    }
                }

                //parse the document
                using (var reader = new StructuredStorageReader(procFile.File.FullName))
                {
                    var xlsDoc = new XlsDocument(reader);

                    var outType = Converter.DetectOutputType(xlsDoc);
                    string conformOutputFile = Converter.GetConformFilename(ChoosenOutputFile, outType);
                    using (var spreadx = SpreadsheetDocument.Create(conformOutputFile, outType))
                    {
                        //start time
                        var start = DateTime.Now;
                        TraceLogger.Info("Converting file {0} into {1}", InputFile, conformOutputFile);

                        Converter.Convert(xlsDoc, spreadx);

                        var end = DateTime.Now;
                        var diff = end.Subtract(start);
                        TraceLogger.Info("Conversion of file {0} finished in {1} seconds", InputFile, diff.TotalSeconds.ToString(CultureInfo.InvariantCulture));
                    }
                }
            }
            catch (ParseException ex)
            {
                TraceLogger.Error("Could not convert {0} because it was created by an unsupported application (Excel 95 or older).", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (DirectoryNotFoundException ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
            catch (FileNotFoundException ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
            catch (ZipCreationException ex)
            {
                TraceLogger.Error("Could not create output file {0}.", ChoosenOutputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (Exception ex)
            {
                TraceLogger.Error("Conversion of file {0} failed.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
        }
    }
}
