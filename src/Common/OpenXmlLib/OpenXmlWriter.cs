using System;
using System.Text;
using System.Xml;
using b2xtranslator.ZipUtils;
using System.IO;

namespace b2xtranslator.OpenXmlLib
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
            if (this._zipOutputStream != null || this._delegateWriter != null)
            {
                this.Close();
            }

            this._zipOutputStream = ZipFactory.CreateArchive(fileName);
        }

        public override void Close()
        {
            // close streams
            if (this._delegateWriter != null)
            {
                this._delegateWriter.Close();
                this._delegateWriter = null;

            }
            if (this._zipOutputStream != null)
            {
                this._zipOutputStream.Close();
                this._zipOutputStream.Dispose();
                this._zipOutputStream = null;
            }
        }

        public void AddPart(string fullName) 
        {
            if (this._delegateWriter != null)
            {
                this._delegateWriter.Close();
                this._delegateWriter = null;
            }

            // the path separator in the package should be a forward slash
            this._zipOutputStream.AddEntry(fullName.Replace('\\', '/'));
        }

        protected XmlWriter XmlWriter
        {
            get
            {
                if (this._delegateWriter == null)
                {
                    this._delegateWriter = XmlWriter.Create(this._zipOutputStream, this._delegateWriterSettings);
                }
                return this._delegateWriter;
            }
        }

        public void WriteRawBytes(byte[] buffer, int index, int count)
        {
            this._zipOutputStream.Write(buffer, index, count);
        }

        public void Write(Stream stream)
        {
            const int blockSize = 4096;
            var buffer = new byte[blockSize];
            int bytesRead;
            while ((bytesRead = stream.Read(buffer, 0, blockSize)) > 0)
            {
                this._zipOutputStream.Write(buffer, 0, bytesRead);
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
