using System;
using System.Text;
using System.Xml;
using System.IO;
using System.IO.Compression; // Replaces using b2xtranslator.ZipUtils;

namespace b2xtranslator.OpenXmlLib
{
    public sealed class OpenXmlWriter : XmlWriter
    {
        /// <summary>Hold the settings required in Open XML ZIP files.</summary>
        readonly static XmlWriterSettings xmlWriterSettings = new XmlWriterSettings
        {
            OmitXmlDeclaration = false,
            CloseOutput = false,
            Encoding = Encoding.UTF8,
            Indent = true,
            ConformanceLevel = ConformanceLevel.Document
        };

        /// <summary>Hold the XML writer to populate the current ZIP entry.</summary>
        XmlWriter xmlEntryWriter;

        /// <summary>Hold an optional file output stream, only populated if opened on a file.</summary>
        FileStream fileOutputStream;

        /// <summary>Holds the ZIP archive the XML is being written to.</summary>
        ZipArchive outputArchive;

        /// <summary>Holds the current ZIP entry, created by <see cref="AddPart"/>.</summary>
        ZipArchiveEntry currentEntry;

        /// <summary>Holds the open stream to write to <see cref="currentEntry"/></summary>
        Stream entryStream;


        public OpenXmlWriter()
        {
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                this.Close();

            base.Dispose(disposing);
        }

        public void Open(string fileName)
        {
            this.Close();
            this.fileOutputStream = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            this.outputArchive = new ZipArchive(this.fileOutputStream, ZipArchiveMode.Update);
        }

        public void Open(Stream output)
        {
            this.Close();
            this.outputArchive = new ZipArchive(output, ZipArchiveMode.Update);
        }

        public override void Close()
        {
            // close streams
            if (this.xmlEntryWriter != null)
            {
                this.xmlEntryWriter.Close();
                this.xmlEntryWriter = null;
            }

            if (this.entryStream != null)
            {
                this.entryStream.Close();
                this.entryStream = null;
            }

            this.currentEntry = null;

            if (this.outputArchive != null)
            {
                this.outputArchive.Dispose();
                this.outputArchive = null;
            }

            if (this.fileOutputStream != null)
            {
                this.fileOutputStream.Close();
                this.fileOutputStream.Dispose();
                this.fileOutputStream = null;
            }
        }

        public void AddPart(string fullName)
        {
            if (this.xmlEntryWriter != null)
            {
                this.xmlEntryWriter.Close();
                this.xmlEntryWriter = null;
            }

            if (this.entryStream != null)
            {
                this.entryStream.Close();
                this.entryStream = null;
            }

            // the path separator in the package should be a forward slash
            this.currentEntry = this.outputArchive.CreateEntry(fullName.Replace('\\', '/'));

            // Create the stream for the current entry
            this.entryStream = this.currentEntry.Open();
        }

        /// <summary>Get or create an XML writer for the current ZIP entry.</summary>
        XmlWriter XmlWriter =>
            this.xmlEntryWriter ?? (this.xmlEntryWriter = XmlWriter.Create(this.entryStream, xmlWriterSettings));

        public void WriteRawBytes(byte[] buffer, int index, int count) =>
            this.entryStream.Write(buffer, index, count);

        public void Write(Stream stream)
        {
            const int blockSize = 4096;
            var buffer = new byte[blockSize];
            int bytesRead;
            while ((bytesRead = stream.Read(buffer, 0, blockSize)) > 0)
                this.entryStream.Write(buffer, 0, bytesRead);
        }

        public override void WriteStartElement(string prefix, string localName, string ns) =>
            this.XmlWriter.WriteStartElement(prefix, localName, ns);

        public override void WriteEndElement() =>
            this.XmlWriter.WriteEndElement();

        public override void WriteStartAttribute(string prefix, string localName, string ns) =>
            this.XmlWriter.WriteStartAttribute(prefix, localName, ns);

        public void WriteAttributeValue(string prefix, string localName, string ns, string value) =>
            this.XmlWriter.WriteAttributeString(prefix, localName, ns, value);

        public override void WriteEndAttribute() =>
            this.XmlWriter.WriteEndAttribute();

        public override void WriteString(string text) =>
            this.XmlWriter.WriteString(text);

        public override void WriteFullEndElement() =>
            this.XmlWriter.WriteFullEndElement();

        public override void WriteCData(string s) =>
            this.XmlWriter.WriteCData(s);

        public override void WriteComment(string s) =>
            this.XmlWriter.WriteComment(s);

        public override void WriteProcessingInstruction(string name, string text) =>
            this.XmlWriter.WriteProcessingInstruction(name, text);

        public override void WriteEntityRef(string name) =>
            this.XmlWriter.WriteEntityRef(name);

        public override void WriteCharEntity(char c) =>
            this.XmlWriter.WriteCharEntity(c);

        public override void WriteWhitespace(string s) =>
            this.XmlWriter.WriteWhitespace(s);

        public override void WriteSurrogateCharEntity(char lowChar, char highChar) =>
            this.XmlWriter.WriteSurrogateCharEntity(lowChar, highChar);

        public override void WriteChars(char[] buffer, int index, int count) =>
            this.XmlWriter.WriteChars(buffer, index, count);

        public override void WriteRaw(char[] buffer, int index, int count) =>
            this.XmlWriter.WriteRaw(buffer, index, count);

        public override void WriteRaw(string data) =>
            this.XmlWriter.WriteRaw(data);

        public override void WriteBase64(byte[] buffer, int index, int count) =>
            this.XmlWriter.WriteBase64(buffer, index, count);

        public override WriteState WriteState =>
            this.XmlWriter.WriteState;

        public override void Flush() =>
            this.XmlWriter.Flush();

        public override string LookupPrefix(string ns) =>
            this.XmlWriter.LookupPrefix(ns);

        public override void WriteDocType(string name, string pubid, string sysid, string subset) =>
            throw new NotImplementedException();

        public override void WriteEndDocument() =>
            this.XmlWriter.WriteEndDocument();

        public override void WriteStartDocument(bool standalone) =>
            this.XmlWriter.WriteStartDocument(standalone);

        public override void WriteStartDocument() =>
            this.XmlWriter.WriteStartDocument();
    }
}