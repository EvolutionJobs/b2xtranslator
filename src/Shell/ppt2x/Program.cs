/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Globalization;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;
using DIaLOGIKa.b2xtranslator.PresentationMLMapping;
using DIaLOGIKa.b2xtranslator.ZipUtils;
using System.Threading;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;
using DIaLOGIKa.b2xtranslator.Shell;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.ppt2x
{
    public class Program : CommandLineTranslator
    {
        public static string ToolName = "ppt2x";
        public static string RevisionResource = "DIaLOGIKa.b2xtranslator.ppt2x.revision.txt";
        public static string ContextMenuInputExtension = ".ppt";
        public static string ContextMenuText = "Convert to .pptx";

        public static void Main(string[] args)
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            ParseArgs(args, ToolName);

            InitializeLogger();

            PrintWelcome(ToolName, RevisionResource);

            try
            {
                if (InputFile.Contains("*.ppt"))
                {

                    var files = Directory.GetFiles(InputFile.Replace("*.ppt", ""), "*.ppt");

                    foreach (var file in files)
                    {
                        if (new FileInfo(file).Extension.ToLower().EndsWith("ppt"))
                        {
                            ChoosenOutputFile = null;
                            processFile(file);
                        }
                    }

                }
                else
                {
                    processFile(InputFile);
                }


            }
            catch (ZipCreationException ex)
            {
                TraceLogger.Error("Could not create output file {0}.", ChoosenOutputFile);
                //TraceLogger.Error("Perhaps the specified outputfile was a directory or contained invalid characters.");
                TraceLogger.Debug(ex.ToString());
            }
            catch (FileNotFoundException ex)
            {
                TraceLogger.Error("Could not read input file {0}.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (MagicNumberException)
            {
                TraceLogger.Error("Input file {0} is not a valid PowerPoint 97-2007 file.", InputFile);
            }
            catch (InvalidStreamException)
            {
                TraceLogger.Error("Input file {0} is not a valid PowerPoint 97-2007 file.", InputFile);
            }
            catch (InvalidRecordException)
            {
                TraceLogger.Error("Input file {0} is not a valid PowerPoint 97-2007 file.", InputFile);
            }
            catch (StreamNotFoundException)
            {
                TraceLogger.Error("Input file {0} is not a valid PowerPoint 97-2007 file.", InputFile);
            }
            catch (Exception ex)
            {
                TraceLogger.Error("Conversion of file {0} failed.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }

            TraceLogger.Info("End of program");
        }

        private static void processFile(string InputFile)
        {
            // copy processing file
            var procFile = new ProcessingFile(InputFile);

            //make output file name
            if (ChoosenOutputFile == null)
            {
                if (InputFile.Contains("."))
                {
                    ChoosenOutputFile = InputFile.Remove(InputFile.LastIndexOf(".")) + ".pptx";
                }
                else
                {
                    ChoosenOutputFile = InputFile + ".pptx";
                }
            }

            //open the reader
            using (var reader = new StructuredStorageReader(procFile.File.FullName))
            {
                // parse the ppt document
                var ppt = new PowerpointDocument(reader);

                // detect document type and name
                var outType = Converter.DetectOutputType(ppt);
                string conformOutputFile = Converter.GetConformFilename(ChoosenOutputFile, outType);

                // create the pptx document
                var pptx = PresentationDocument.Create(conformOutputFile, outType);

                //start time
                var start = DateTime.Now;
                TraceLogger.Info("Converting file {0} into {1}", InputFile, conformOutputFile);

                // convert
                Converter.Convert(ppt, pptx);

                // stop time
                var end = DateTime.Now;
                var diff = end.Subtract(start);
                TraceLogger.Info("Conversion of file {0} finished in {1} seconds", InputFile, diff.TotalSeconds.ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}