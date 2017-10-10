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
using System.Text;
using System.Xml;
using DIaLOGIKa.b2xtranslator.ZipUtils;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib
{
    public class OpenXmlWriter : XmlWriter
    {
        protected XmlWriter _delegateWriter;
        protected XmlWriterSettings _delegateWriterSettings;

        protected ZipWriter _zipOutputStream;
        
        public OpenXmlWriter()
        {
            var _delegateWriterSettings = new XmlWriterSettings
            {
                OmitXmlDeclaration = false,
                CloseOutput = false,
                Encoding = Encoding.UTF8,
                Indent = true,
                ConformanceLevel = ConformanceLevel.Document
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                this.Close();
            }
            base.Dispose(disposing);

        }

        public void Open(string fileName)
        {
            if (_zipOutputStream != null || _delegateWriter != null)
            {
                this.Close();
            }

            _zipOutputStream = ZipFactory.CreateArchive(fileName);
        }

        public override void Close()
        {
            // close streams
            if (_delegateWriter != null)
            {
                _delegateWriter.Close();
                _delegateWriter = null;

            }
            if (_zipOutputStream != null)
            {
                _zipOutputStream.Close();
                _zipOutputStream.Dispose();
                _zipOutputStream = null;
            }
        }

        public void AddPart(string fullName) 
        {
            if (_delegateWriter != null)
            {
                _delegateWriter.Close();
                _delegateWriter = null;
            }

            // the path separator in the package should be a forward slash
            _zipOutputStream.AddEntry(fullName.Replace('\\', '/'));
        }

        protected XmlWriter XmlWriter
        {
            get
            {
                if (_delegateWriter == null)
                {
                    _delegateWriter = XmlWriter.Create(_zipOutputStream, _delegateWriterSettings);
                }
                return _delegateWriter;
            }
        }

        public void WriteRawBytes(byte[] buffer, int index, int count)
        {
            _zipOutputStream.Write(buffer, index, count);
        }

        public void Write(Stream stream)
        {
            const int blockSize = 4096;
            var buffer = new byte[blockSize];
            int bytesRead;
            while ((bytesRead = stream.Read(buffer, 0, blockSize)) > 0)
            {
                _zipOutputStream.Write(buffer, 0, bytesRead);
            }
        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            this.XmlWriter.WriteStartElement(prefix, localName, ns);
        }

        public override void WriteEndElement()
        {
            this.XmlWriter.WriteEndElement();
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            this.XmlWriter.WriteStartAttribute(prefix, localName, ns);
        }

        public void WriteAttributeValue(string prefix, string localName, string ns, string value)
        {
            this.XmlWriter.WriteAttributeString(prefix, localName, ns, value);
        }

        public override void WriteEndAttribute()
        {
            this.XmlWriter.WriteEndAttribute();
        }

        public override void WriteString(string text)
        {
            this.XmlWriter.WriteString(text);
        }

        public override void WriteFullEndElement()
        {
            this.XmlWriter.WriteFullEndElement();
        }

        public override void WriteCData(string s)
        {
            this.XmlWriter.WriteCData(s);
        }

        public override void WriteComment(string s)
        {
            this.XmlWriter.WriteComment(s);
        }

        public override void WriteProcessingInstruction(string name, string text)
        {
            this.XmlWriter.WriteProcessingInstruction(name, text);
        }

        public override void WriteEntityRef(string name)
        {
            this.XmlWriter.WriteEntityRef(name);
        }

        public override void WriteCharEntity(char c)
        {
            this.XmlWriter.WriteCharEntity(c);
        }

        public override void WriteWhitespace(string s)
        {
            this.XmlWriter.WriteWhitespace(s);
        }

        public override void WriteSurrogateCharEntity(char lowChar, char highChar)
        {
            this.XmlWriter.WriteSurrogateCharEntity(lowChar, highChar);
        }

        public override void WriteChars(char[] buffer, int index, int count)
        {
            this.XmlWriter.WriteChars(buffer, index, count);
        }

        public override void WriteRaw(char[] buffer, int index, int count)
        {
            this.XmlWriter.WriteRaw(buffer, index, count);
        }

        public override void WriteRaw(string data)
        {
            this.XmlWriter.WriteRaw(data);
        }

        public override void WriteBase64(byte[] buffer, int index, int count)
        {
            this.XmlWriter.WriteBase64(buffer, index, count);
        }

        public override WriteState WriteState
        {
            get
            {
                return this.XmlWriter.WriteState;
            }
        }

        public override void Flush()
        {
            this.XmlWriter.Flush();
        }

        public override string LookupPrefix(string ns)
        {
            return this.XmlWriter.LookupPrefix(ns);
        }

        public override void WriteDocType(string name, string pubid, string sysid, string subset)
        {
            throw new NotImplementedException();
        }

        public override void WriteEndDocument()
        {
            this.XmlWriter.WriteEndDocument();
        }

        public override void WriteStartDocument(bool standalone)
        {
            this.XmlWriter.WriteStartDocument(standalone);
        }

        public override void WriteStartDocument()
        {
            this.XmlWriter.WriteStartDocument();
        }
    }
}
