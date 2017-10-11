using b2xtranslator.OfficeDrawing;
using b2xtranslator.OpenXmlLib.PresentationML;
using b2xtranslator.PptFileFormat;
using b2xtranslator.PresentationMLMapping;
using b2xtranslator.Shell;
using b2xtranslator.StructuredStorage.Common;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using System;
using System.Globalization;
using System.IO;
using System.Threading;

namespace b2xtranslator.ppt2x
{
    public class Program : CommandLineTranslator
    {
        public static string ToolName = "ppt2x";
        public static string RevisionResource = "b2xtranslator.ppt2x.revision.txt";
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

                    foreach (string file in files)
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