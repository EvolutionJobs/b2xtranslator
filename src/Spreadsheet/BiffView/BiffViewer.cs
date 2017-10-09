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
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Collections.Generic;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.BiffView
{
    public class BiffViewer
    {
        struct BiffHeader
        {
            public RecordType id;
            public UInt16 length;
        }

        private BiffViewerOptions _options;
        private BackgroundWorker _backgroundWorker;
        private bool _isCancelled = false;

        public BiffViewerOptions Options
        {
            get { return _options; }
            set { _options = value; }
        }

        public BiffViewer(BiffViewerOptions options)
        {
            this._options = options;
        }

        public void DoTheMagic(BackgroundWorker backgroundWorker)
        {
            _backgroundWorker = backgroundWorker;
            DoTheMagic();
            _backgroundWorker = null;
        }

        public void DoTheMagic()
        {
            try
            {
                using (StructuredStorageReader reader = new StructuredStorageReader(this.Options.InputDocument))
                {
                    IStreamReader workbookReader = new VirtualStreamReader(reader.GetStream("Workbook"));

                    using (StreamWriter sw = this.Options.Mode == BiffViewerMode.File 
                        ? File.CreateText(this.Options.OutputFileName) 
                        : new StreamWriter(Console.OpenStandardOutput()))
                    {
                        sw.AutoFlush = true;

                        if (this.Options.PrintTextOnly)
                        {
                            PrintText(sw, workbookReader);
                        }
                        else
                        {
                            PrintHtml(sw, workbookReader);
                        }

                        if (!_isCancelled && this.Options.ShowInBrowser && this.Options.Mode == BiffViewerMode.File)
                        {
                            Util.VisitLink(this.Options.OutputFileName);
                        }
                    }
                }
            }
            catch (MagicNumberException ex)
            {
                if (this.Options.ShowErrors)
                {
                    MessageBox.Show(string.Format("This file is not a valid Excel file ({0})", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    Console.WriteLine(ex.ToString());
                }
            }
            catch (Exception ex)
            {
                if (this.Options.ShowErrors)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }

        protected void ShowInBrowser()
        {
            Util.VisitLink(this.Options.OutputFileName);
        }

        protected void PrintText(StreamWriter sw, IStreamReader workbookReader)
        {
            BiffHeader bh;

            try
            {
                while (workbookReader.BaseStream.Position < workbookReader.BaseStream.Length)
                {
                    bh.id = (RecordType)workbookReader.ReadUInt16();
                    bh.length = workbookReader.ReadUInt16();

                    byte[] buffer = new byte[bh.length];
                    if (bh.length != workbookReader.Read(buffer, bh.length))
                        sw.WriteLine("EOF");

                    sw.Write("BIFF {0}\t{1}\t", bh.id, bh.length);
                    //Dump(buffer);
                    int count = 0;
                    foreach (byte b in buffer)
                    {
                        sw.Write("{0:X02} ", b);
                        count++;
                        if (count % 16 == 0 && count < buffer.Length)
                        {
                            sw.Write("\n\t\t\t");
                        }
                    }
                    sw.Write("\n");

                    if (_backgroundWorker != null)
                    {
                        int progress = 100;

                        if (sw.BaseStream.Length != 0)
                        {
                            progress = (int)(100 * workbookReader.BaseStream.Position / workbookReader.BaseStream.Length);
                        }
                        _backgroundWorker.ReportProgress(progress);

                        if (_backgroundWorker.CancellationPending)
                        {
                            _isCancelled = true;
                            break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        protected void PrintHtml(StreamWriter sw, IStreamReader workbookReader)
        {
            BiffHeader bh = new BiffHeader();
            BiffHeader prevHeader;
            Stack<BiffHeader> blocks = new Stack<BiffHeader>();

            int indentLevel = 0;

            Uri baseUrl = new Uri(new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName);

            sw.WriteLine("<html>");
            sw.WriteLine("<head>");
            sw.WriteLine("<title>" + this.Options.InputDocument + "</title>");
            sw.WriteLine("<link href=\"style.css\" rel=\"stylesheet\" type=\"text/css\">");
            sw.WriteLine("<style>");
            sw.WriteLine("    td { font-family: Monospace, Courier; vertical-align: top; border-top: 1px solid black; padding-left: 2px; padding-right: 2px }");
            sw.WriteLine("    table { border: 1px solid black; empty-cells:show; border-collapse:collapse}");
            sw.WriteLine("</style>");
            sw.WriteLine("</head>");
            sw.WriteLine("<body>");
            sw.WriteLine("<table>");

            try
            {
                while (workbookReader.BaseStream.Position < workbookReader.BaseStream.Length)
                {
                    long offset = workbookReader.BaseStream.Position;

                    prevHeader = bh;

                    bh.id = (RecordType)workbookReader.ReadUInt16();
                    bh.length = workbookReader.ReadUInt16();

                    // check if the record is the BOF record 
                    string documentType = "";
                    if (bh.id == RecordType.BOF)
                    {
                        BOF bof = new BOF(workbookReader, (XlsFileFormat.RecordType)bh.id, bh.length);
                        documentType = bof.docType.ToString();
                        // seek back
                        workbookReader.BaseStream.Seek(-bh.length, System.IO.SeekOrigin.Current);
                    }

                    string strId = ((XlsFileFormat.RecordType)bh.id).ToString();
                    string strCurrentBlock = "";
                    
                    int indent = 4 * indentLevel;
                            
                    if (bh.id == RecordType.Begin)
                        //|| bh.id == RecordType.StartObject
                        //|| bh.id == RecordType.StartBlock)
                    {
                        indent += 2; 
                        indentLevel++;
                        blocks.Push(prevHeader);
                        strCurrentBlock = " " + prevHeader.id;
                    }
                    else if (bh.id == RecordType.End)
                        //|| bh.id == RecordType.EndObject
                        //|| bh.id == RecordType.EndBlock)
                    {
                        indent -= 2;
                        strCurrentBlock = " " + blocks.Pop().id;
                    }
                    
                    sw.WriteLine("<tr>");
                    {
                        byte[] buffer = workbookReader.ReadBytes(bh.length);

                        sw.WriteLine("<td>");
                        {
                            string url = string.Format("{0}/xlsspec/{1}.html", baseUrl, strId);
                            Uri uri = new Uri(url);
                            if (!File.Exists(uri.LocalPath))
                            {
                                // unspecified record id
                                url = string.Format("{0}/xlsspec/404.html", baseUrl);
                            }

                            // write record type
                            sw.Write("BIFF ");
                            for (int i = 0; i < indent; i++)
                            {
                                sw.Write("&nbsp;");
                            }
                            sw.WriteLine("<a href=\"{0}\">{1}</a> {2}{3} (0x{4:X02})", url, strId, strCurrentBlock, documentType, (int)bh.id);
                        }
                        sw.WriteLine("</td>");

                        // offset
                        sw.WriteLine("<td>");
                        {
                            sw.WriteLine("0x{0:X04}", offset);
                        }
                        sw.WriteLine("</td>");

                        // record length
                        sw.WriteLine("<td>");
                        {
                            sw.WriteLine("0x{0:X02}", bh.length);
                        }
                        sw.WriteLine("</td>");

                        // raw data
                        sw.WriteLine("<td>");
                        {
                            //Dump(buffer);
                            int count = 0;
                            foreach (byte b in buffer)
                            {
                                sw.Write("{0:X02}&nbsp;", b);
                                count++;
                                if (count % 16 == 0 && count < buffer.Length)
                                    sw.Write("</br>");
                                else if (count % 8 == 0 && count < buffer.Length)
                                    sw.Write("&nbsp;");
                            }
                        }
                        sw.Write("</td>");
                    }
                    sw.Write("</tr>");

                    if (bh.id == RecordType.End)
                        //|| bh.id == RecordType.EndObject
                        //|| bh.id == RecordType.EndBlock)
                    {
                        indentLevel--;
                    }

                    if (_backgroundWorker != null)
                    {
                        int progress = 100;

                        if (sw.BaseStream.Length != 0)
                        {
                            progress = (int)(100 * workbookReader.BaseStream.Position / workbookReader.BaseStream.Length);
                        }
                        _backgroundWorker.ReportProgress(progress);

                        if (_backgroundWorker.CancellationPending)
                        {
                            _isCancelled = true;
                            break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            sw.WriteLine("</table>");
            sw.WriteLine("</body>");
            sw.WriteLine("</html>");
        }
    }
}
