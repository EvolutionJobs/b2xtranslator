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
using System.IO;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// Extracts the FileSummaryStream 
    /// </summary>
    public class FileInformationExtractor
    {
        public VirtualStream summaryStream;         // Summary stream 
        public VirtualStreamReader SummaryStream;         // Summary stream 

        public string Title;

        public string buffer; 

        struct BiffHeader
        {
            public RecordType id;
            public UInt16 length;
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="sum">Summary stream </param>
        public FileInformationExtractor(VirtualStream sum)
        {
            this.Title = null; 
            if (sum == null)
            {
                throw new ExtractorException(ExtractorException.NULLPOINTEREXCEPTION); 
            }
            this.summaryStream = sum; 
            this.SummaryStream = new VirtualStreamReader(sum);
            this.extractData(); 


        }

        /// <summary>
        /// Extracts the data from the stream 
        /// </summary>
        public void extractData()
        {
            BiffHeader bh;
            StreamWriter sw = null;
            sw = new StreamWriter(Console.OpenStandardOutput());
            try
            {
                while (this.SummaryStream.BaseStream.Position < this.SummaryStream.BaseStream.Length)
                {
                    bh.id = (RecordType)this.SummaryStream.ReadUInt16();
                    bh.length = this.SummaryStream.ReadUInt16();

                    byte[] buf = new byte[bh.length];
                    if (bh.length != this.SummaryStream.ReadByte())
                        sw.WriteLine("EOF");

                    sw.Write("BIFF {0}\t{1}\t", bh.id, bh.length);
                    //Dump(buffer);
                    int count = 0;
                    foreach (byte b in buf)
                    {
                        sw.Write("{0:X02} ", b);
                        count++;
                        if (count % 16 == 0 && count < buf.Length)
                            sw.Write("\n\t\t\t");
                    }
                    sw.Write("\n");
                }

            }
            catch (Exception ex)
            {
                TraceLogger.Error(ex.Message);
                TraceLogger.Debug(ex.ToString());
            }
            this.buffer = sw.ToString();
         }

        /// <summary>
        /// A normal overload ToString Method 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            string returnvalue = "Title: " + this.Title;
            return returnvalue; 
        }
    }
}
