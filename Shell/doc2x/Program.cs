using b2xtranslator.DocFileFormat;
using b2xtranslator.OpenXmlLib.WordprocessingML;
using b2xtranslator.Shell;
using b2xtranslator.StructuredStorage.Common;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using b2xtranslator.WordprocessingMLMapping;
using System;
using System.Globalization;
using System.IO;

namespace b2xtranslator.doc2x
{
    public class Program : CommandLineTranslator
    {
        public static string ToolName = "doc2x";
        public static string RevisionResource = "b2xtranslator.doc2x.revision.txt";
        public static string ContextMenuInputExtension = ".doc";
        public static string ContextMenuText = "Convert to .docx";

        public static void Main(string[] args)
        {
            ParseArgs(args, ToolName);

            InitializeLogger();

            PrintWelcome(ToolName, RevisionResource);

            // convert
            try
            {
                //copy processing file
                var procFile = new ProcessingFile(InputFile);

                //make output file name
                if (ChoosenOutputFile == null)
                {
                    if (InputFile.Contains("."))
                    {
                        ChoosenOutputFile = InputFile.Remove(InputFile.LastIndexOf(".")) + ".docx";
                    }
                    else
                    {
                        ChoosenOutputFile = InputFile + ".docx";
                    }
                }

                //open the reader
                using (var reader = new StructuredStorageReader(procFile.File.FullName))
                {
                    //parse the input document
                    var doc = new WordDocument(reader);

                    //prepare the output document
                    var outType = Converter.DetectOutputType(doc);
                    string conformOutputFile = Converter.GetConformFilename(ChoosenOutputFile, outType);
                    var docx = WordprocessingDocument.Create(conformOutputFile, outType);

                    //start time
                    var start = DateTime.Now;
                    TraceLogger.Info("Converting file {0} into {1}", InputFile, conformOutputFile);

                    //convert the document
                    Converter.Convert(doc, docx);

                    var end = DateTime.Now;
                    var diff = end.Subtract(start);
                    TraceLogger.Info("Conversion of file {0} finished in {1} seconds", InputFile, diff.TotalSeconds.ToString(CultureInfo.InvariantCulture));
                }
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
            catch (ReadBytesAmountMismatchException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (MagicNumberException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (UnspportedFileVersionException ex)
            {
                TraceLogger.Error("File {0} has been created with a Word version older than Word 97.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (ByteParseException ex)
            {
                TraceLogger.Error("Input file {0} is not a valid Microsoft Word 97-2003 file.", InputFile);
                TraceLogger.Debug(ex.ToString());
            }
            catch (MappingException ex)
            {
                TraceLogger.Error("There was an error while converting file {0}: {1}", InputFile, ex.Message);
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
