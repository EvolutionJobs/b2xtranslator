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
using System.Windows.Forms;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.BiffView
{
    public class CommandLineException : Exception
    {
        public CommandLineException(string message)
            : base(message)
        {
        }
    }

    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            BiffViewerOptions options = ParseCommandLine(args);

            if (options != null)
            {
                if (args.Length == 0)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new MainWindow());
                }
                else
                {
                    try
                    {
                        BiffViewer viewer = new BiffViewer(options);
                        viewer.DoTheMagic();
                    }
                    catch (CommandLineException cex)
                    {
                        Console.WriteLine(cex.Message);
                        PrintUsage();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("An unexpected error occurred.");
                        Console.WriteLine(ex.ToString());
                    }
                }
            }
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Usage: BiffViewer.exe [OPTIONS] [FILE]");
            Console.WriteLine("Where options are:");
            Console.WriteLine("    -o, --out outputFile     Prints the list of BIFF records to file outputFile.");
            Console.WriteLine("    -s, --show               Opens the output in the default browser.");
            Console.WriteLine("    -t, --text               Outputs text only.");
            Console.WriteLine("    -r, --range start [end]  Shows only BIFF records within the range from 'start' to 'end'.");
            Console.WriteLine("                             The parameters are integer numbers. If 'end' is omitted, all BIFF records");
            Console.WriteLine("                             to the end of the document are printed.");
            Console.WriteLine("    -c, --console            Does not display an interactive dialog.");
            Console.WriteLine("    -v, --version            Display the application version.");
            Console.WriteLine("    -h, --help               Display this help message.");
        }

        private static BiffViewerOptions ParseCommandLine(string[] args)
        {
            BiffViewerOptions options = new BiffViewerOptions();
            options.UseTempFolder = false;

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLowerInvariant())
                {
                    case "-o":
                    case "/o":
                    case "--out":
                        if (++i == args.Length)
                        {
                            throw new CommandLineException("Output file not specified.");
                        }
                        options.OutputFileName = args[i];
                        break;
                    case "-s":
                    case "/s":
                    case "--show":
                        options.ShowInBrowser = true;
                        break;
                    case "-t":
                    case "/t":
                    case "--text":
                        options.PrintTextOnly = true;
                        break;
                    case "-r":
                    case "/r":
                    case "--range":
                        if (++i == args.Length)
                        {
                            throw new CommandLineException("Start position for range not specified.");
                        }
                        long dummy;
                        if (long.TryParse(args[i], out dummy))
                        {
                            options.StartPosition = dummy;
                        }
                        
                        if (i + 1 != args.Length)
                        {
                            if (long.TryParse(args[i + 1], out dummy))
                            {
                                options.EndPosition = dummy;
                            }
                        }
                        break;
                    case "-v":
                    case "/v":
                    case "--version":
                        Console.WriteLine("{0} {1}", Util.AssemblyProduct, Util.AssemblyVersion);
                        return null;
                    case "-h":
                    case "/h":
                    case "/?":
                    case "--help":
                        PrintUsage();
                        return null;
                    default:
                        if (!String.IsNullOrEmpty(options.InputDocument))
                        {
                            throw new CommandLineException("Wrong command line parameters.");
                        }
                        options.InputDocument = args[i];
                        break;
                }
            }
            return options;
        }
    }
}