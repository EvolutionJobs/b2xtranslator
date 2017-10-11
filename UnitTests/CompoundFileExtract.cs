using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace UnitTests
{


    /// <summary>
    /// Test application which extracts streams from compound files.
    /// Author: math
    /// </summary>
    public class CompoundFileExtract
    {
        static void RunCompoundTests(string[] args)
        {            
            const int bytesToReadAtOnce = 1024;
            var invalidChars = Path.GetInvalidFileNameChars();

            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Error;
            //var consoleTracer = new ConsoleTraceListener();
            //Trace.Listeners.Add(consoleTracer);
            Trace.AutoFlush = true;

            if (args.Length < 1)
            {
                Console.WriteLine("No parameter found. Please specify one or more compound document file(s).");
                return;
            }

            foreach (string file in args)
            {

                StructuredStorageReader storageReader = null;
                var begin = DateTime.Now;
                var extractionTime = new TimeSpan();

                try
                {                   
                    // init StorageReader
                    storageReader = new StructuredStorageReader(file);

                    // read stream _entries
                    var streamEntries = storageReader.AllStreamEntries;

                    // create valid path names
                    var PathNames = new Dictionary<string, string>();
                    foreach (var entry in streamEntries)
                    {
                        string name = entry.Path;
                        for (int i = 0; i < invalidChars.Length; i++)
                        {
                            name = name.Replace(invalidChars[i], '_');
                        }
                        PathNames.Add(entry.Path, name);
                    }

                    // create output directory
                    string outputDir = '_' + file.Replace('.', '_');
                    outputDir = outputDir.Replace(':', '_'); 
                    Directory.CreateDirectory(outputDir);

                    // for each stream                    
                    foreach (string key in PathNames.Keys)
                    {
                        // get virtual stream by path name
                        IStreamReader streamReader = new VirtualStreamReader(storageReader.GetStream(key));

                        // read bytes from stream, write them back to disk
                        var fs = new FileStream(outputDir + "\\" + PathNames[key] + ".stream", FileMode.Create);
                        var writer = new BinaryWriter(fs);
                        var array = new byte[bytesToReadAtOnce];
                        int bytesRead;
                        do
                        {
                            bytesRead = streamReader.Read(array);
                            writer.Write(array, 0, bytesRead);
                            writer.Flush();
                        } while (bytesRead == array.Length);
                        
                        writer.Close();
                        fs.Close();
                    }

                    // close storage
                    storageReader.Close();
                    storageReader = null;

                    extractionTime = DateTime.Now - begin;                                        
                    Console.WriteLine("Streams extracted in " + string.Format("{0:N2}", extractionTime.TotalSeconds) + "s. (File: " + file + ")");
                }
                catch (Exception e)
                {
                    Console.WriteLine("*** Error: " + e.Message + " (File: " + file + ")");
                }
                finally
                {
                    if (storageReader != null)
                    {
                        storageReader.Close();
                    }
                }
            }
        }
    }
}
