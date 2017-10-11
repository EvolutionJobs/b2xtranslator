using b2xtranslator.OpenXmlLib.SpreadsheetML;
using b2xtranslator.Shell;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.SpreadsheetMLMapping;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using System;
using System.Globalization;
using System.IO;

namespace b2xtranslator.xls2x
{
    class Program : CommandLineTranslator
    {
        public static string ToolName = "xls2x";
        public static string RevisionResource = "b2xtranslator.xls2x.revision.txt";
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
            catch (Exception ex)
            {
                TraceLogger.Error("Conversion of file {0} failed.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
        }
    }
}
