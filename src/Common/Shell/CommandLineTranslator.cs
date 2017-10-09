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
using System.Linq;
using System.Text;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Tools;
using System.Reflection;
using System.IO;
using Microsoft.Win32;

namespace DIaLOGIKa.b2xtranslator.Shell
{
    public class CommandLineTranslator
    {
        // parsed arguments
        public static string InputFile;
        public static string ChoosenOutputFile;
        public static bool CreateContextMenuEntry;

        public static void InitializeLogger()
        {
            // let the Console listen to the Trace messages
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
        }

        /// <summary>
        /// Prints the heading row of the tool
        /// </summary>
        public static void PrintWelcome(string toolname, string revisionResource)
        {
            bool backup = TraceLogger.EnableTimeStamp;
            TraceLogger.EnableTimeStamp = false;
            StringBuilder welcome = new StringBuilder();
            welcome.Append("Welcome to ");
            welcome.Append(toolname);
            welcome.Append(" (r");
            welcome.Append(getRevision(revisionResource));
            welcome.Append(")\n");
            welcome.Append("Copyright (c) 2009, DIaLOGIKa. All rights reserved.");
            welcome.Append("\n");
            TraceLogger.Simple(welcome.ToString());
            TraceLogger.EnableTimeStamp = backup;
        }

        /// <summary>
        /// Prints the usage row of the tool
        /// </summary>
        public static void PrintUsage(string toolname)
        {
            StringBuilder usage = new StringBuilder();
            usage.AppendLine("Usage: " + toolname + " [-c | inputfile] [-o outputfile] [-v level] [-?]");
            usage.AppendLine("-o <outputfile>  change output filename");
            usage.AppendLine("-v <level>     set trace level, where <level> is one of the following:");
            usage.AppendLine("               none (0)    print nothing");
            usage.AppendLine("               error (1)   print all errors");
            usage.AppendLine("               warning (2) print all errors and warnings");
            usage.AppendLine("               info (3)    print all errors, warnings and infos (default)");
            usage.AppendLine("               debug (4)   print all errors, warnings, infos and debug messages");
            usage.AppendLine("-c             create an entry in context menu");
            usage.AppendLine("-?             print this help");
            Console.WriteLine(usage.ToString());
        }

        /// <summary>
        /// Returns the revision that is stored in the embedded resource "revision.txt".
        /// Returns -1 if something goes wrong
        /// </summary>
        /// <returns></returns>
        private static int getRevision(string revisionResource)
        {
            int rev = -1;

            try
            {
                Assembly a = Assembly.GetEntryAssembly();
                Stream s = a.GetManifestResourceStream(revisionResource);
                StreamReader reader = new StreamReader(s);
                rev = Int32.Parse(reader.ReadLine());
                s.Close();
            }
            catch (Exception) { }

            return rev;
        }

        public static RegistryKey GetContextMenuKey(string triggerExtension, string contextMenuText)
        {
            RegistryKey result = null;
            try
            {
                string defaultWord = (string)Registry.ClassesRoot.OpenSubKey(triggerExtension).GetValue("");
                result = Registry.ClassesRoot.CreateSubKey(defaultWord).CreateSubKey("shell").CreateSubKey(contextMenuText);
            }
            catch (Exception)
            {
            }
            return result;
        }

        public static void RegisterForContextMenu(RegistryKey entryKey)
        {
            if (entryKey != null)
            {
                RegistryKey convertCommand = entryKey.CreateSubKey("Command");
                convertCommand.SetValue("", String.Format("\"{0}\" \"%1\"", Assembly.GetCallingAssembly().Location));
            }
        }

        /// <summary>
        /// Parses the arguments of the tool
        /// </summary>
        /// <param name="args">The args array</param>
        public static void ParseArgs(string[] args, string toolName)
        {
            try
            {
                if (args[0] == "-?")
                {
                    PrintUsage(toolName);
                    Environment.Exit(0);
                }
                else if (args[0].ToLower() == "-c")
                {
                    CreateContextMenuEntry = true;
                }
                else
                {
                    InputFile = args[0];
                }

                for (int i = 1; i < args.Length; i++)
                {
                    if (args[i].ToLower() == "-v")
                    {
                        //parse verbose level
                        string verbose = args[i + 1].ToLower();
                        int vLvl;
                        if (Int32.TryParse(verbose, out vLvl))
                        {
                            TraceLogger.LogLevel = (TraceLogger.LoggingLevel)vLvl;
                        }
                        else if (verbose == "error")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Error;
                        }
                        else if (verbose == "warning")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Warning;
                        }
                        else if (verbose == "info")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Info;
                        }
                        else if (verbose == "debug")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Debug;
                        }
                        else if (verbose == "none")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.None;
                        }
                    }
                    else if (args[i].ToLower() == "-o")
                    {
                        //parse output file name
                        ChoosenOutputFile = args[i + 1];
                    }
                }
            }
            catch (Exception)
            {
                TraceLogger.Error("At least one of the required arguments was not correctly set.\n");
                PrintUsage(toolName);
                Environment.Exit(1);
            }
        }
    }
}
