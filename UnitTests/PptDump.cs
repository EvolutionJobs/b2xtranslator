using b2xtranslator.OfficeDrawing;
using b2xtranslator.PptFileFormat;
using b2xtranslator.Shell;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using System;
using System.IO;

namespace UnitTests
{
    public class PptDump
    {
        static void RunPptTest(string inputFile)
        {
            TraceLogger.LogLevel = TraceLogger.LoggingLevel.DebugInternal;

            const string outputDir = "dumps";

            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, true);

            Directory.CreateDirectory(outputDir);

            var procFile = new ProcessingFile(inputFile);

            var file = new StructuredStorageReader(procFile.File.FullName);
            var pptDoc = new PowerpointDocument(file);

            // Dump unknown records
            foreach (var record in pptDoc)
            {
                if (record is UnknownRecord)
                {
                    string filename = string.Format(@"{0}\{1}.record", outputDir, record.GetIdentifier());

                    using (var fs = new FileStream(filename, FileMode.Create))
                    {
                        record.DumpToStream(fs);
                    }
                }
            }

            // Output record tree
            Console.WriteLine(pptDoc);
            Console.WriteLine();

            // Let's make development as easy as pie.
            System.Diagnostics.Debugger.Break();
        }
    }
}
