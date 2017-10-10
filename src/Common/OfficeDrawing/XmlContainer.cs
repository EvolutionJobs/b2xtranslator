

using System;
using System.IO;
using DIaLOGIKa.b2xtranslator.ZipUtils;
using System.Xml;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    /// <summary>
    /// XML containers are containers with a zipped XML payload.
    /// </summary>
    public class XmlContainer : XmlRecord
    {
        public XmlContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            // Note: XmlContainers contain the data of a partial "unfinished"
            // OOXML file (.zip based) as their body.
            //
            // I really don't like writing the data to a temp file just to
            // be able to open it via ZipUtils.
            //
            // Possible alternatives:
            // 1) Using System.IO.Compression -- supports inflation, but can't parse Zip header data
            // 2) Modifying zlib + minizlib + ZipLib so I can pass in bytes, possible, but not worth the effort            

            string tempPath = Path.GetTempFileName();

            try
            {
                using (var fs = new FileStream(tempPath, FileMode.Create))
                {
                    using (var tempStream = new BinaryWriter(fs))
                    {
                        int count = (int)this.Reader.BaseStream.Length;
                        var bytes = this.Reader.ReadBytes(count);

                        tempStream.Write(bytes);

                        tempStream.Flush();
                        fs.Flush();

                        tempStream.Close();
                        fs.Close();
                    }
                }

                using (var zipReader = ZipFactory.OpenArchive(tempPath))
                {
                    this.XmlDocumentElement = ExtractDocumentElement(zipReader, GetRelations(zipReader, ""));
                }
            }
            finally
            {
                try
                {
                    File.Delete(tempPath);
                }
                catch (IOException) { /* OK */ }
            }
        }

        /// <summary>
        /// Get the relations for the specified part.
        /// </summary>
        /// <param name="zipReader">ZipReader for reading from the OOXML package</param>
        /// <param name="forPartPath">Part for which to get relations</param>
        /// <returns>List of Relationship nodes belonging to forFile</returns>
        protected static XmlNodeList GetRelations(ZipReader zipReader, string forPartPath)
        {
            string relPath = GetRelationPath(forPartPath);
            var relStream = zipReader.GetEntry(relPath);

            var relDocument = new XmlDocument();
            relDocument.Load(relStream);

            var rels = relDocument["Relationships"].GetElementsByTagName("Relationship");
            return rels;
        }

        /// <summary>
        /// Get the relation path for the specified part.
        /// </summary>
        /// <param name="forPartPath">Part for which to get relations</param>
        /// <returns>Relation path</returns>
        protected static string GetRelationPath(string forPartPath)
        {
            var directoryPath = "";
            var filePath = "";

            if (forPartPath.Length > 0)
            {
                directoryPath = Path.GetDirectoryName(forPartPath).Replace("\\", "/") + "/";
                filePath = Path.GetFileName(forPartPath);
            }

            string relPath = string.Format("{0}_rels/{1}.rels", directoryPath, filePath);
            return relPath;
        }

        /// <summary>
        /// Method that extracts the actual XmlElement that will be used as this XmlContainer's
        /// XmlDocumentElement based on the relations and a ZipReader for the OOXML package.
        /// 
        /// The default implementation simply returns the root of the first referenced part if
        /// there is only one part.
        /// 
        /// Override this in subclasses to implement behaviour for more complex cases.
        /// </summary>
        /// <param name="zipReader">ZipReader for reading from the OOXML package</param>
        /// <param name="rels">List of Relationship nodes belonging to root part</param>
        /// <returns>The XmlElement that will become this record's XmlDocumentElement</returns>
        protected virtual XmlElement ExtractDocumentElement(ZipReader zipReader, XmlNodeList rels)
        {
            if (rels.Count != 1)
                throw new Exception("Expected actly one Relationship in XmlContainer OOXML doc");

            string partPath = rels[0].Attributes["Target"].Value;
            var partStream = zipReader.GetEntry(partPath);

            var partDoc = new XmlDocument();
            partDoc.Load(partStream);

            return partDoc.DocumentElement;
        }
    }
}