using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Shell;

namespace DIaLOGIKa.b2xtranslator.PptDump
{
    class Program
    {
        static void Main(string[] args)
        {
            TraceLogger.LogLevel = TraceLogger.LoggingLevel.DebugInternal;

            const string outputDir = "dumps";

            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, true);

            Directory.CreateDirectory(outputDir);

            string inputFile = args[0];
            ProcessingFile procFile = new ProcessingFile(inputFile);

            StructuredStorageReader file = new StructuredStorageReader(procFile.File.FullName);
            PowerpointDocument pptDoc = new PowerpointDocument(file);

            // Dump unknown records
            foreach (Record record in pptDoc)
            {
                if (record is UnknownRecord)
                {
                    string filename = String.Format(@"{0}\{1}.record", outputDir, record.GetIdentifier());

                    using (FileStream fs = new FileStream(filename, FileMode.Create))
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
