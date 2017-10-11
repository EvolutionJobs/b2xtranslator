using b2xtranslator.OfficeDrawing;
using b2xtranslator.OpenXmlLib;
using System;
using System.IO;
using System.Xml;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1038)]
    public class Theme : XmlContainer
    {
        public Theme(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {}

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
        /// <param name="rootRels">List of Relationship nodes belonging to root part</param>
        /// <returns>The XmlElement that will become this record's XmlDocumentElement</returns>
        protected override XmlElement ExtractDocumentElement(IZipReader zipReader, XmlNodeList rootRels)
        {
            if (rootRels.Count != 1)
                throw new Exception("Expected actly one Relationship in Theme OOXML doc");

            string managerPath = rootRels[0].Attributes["Target"].Value;
            string managerDirectory = Path.GetDirectoryName(managerPath).Replace("\\", "/");
            XmlNodeList managerRels;

            try
            {
                managerRels = GetRelations(zipReader, managerPath);
            }
            catch (Exception)
            {
                this.XmlDocumentElement = null;
                return null;
            }
           
            if (managerRels.Count != 1)
                throw new Exception("Expected actly one Relationship for Theme manager");

            string partPath = string.Format("{0}/{1}", managerDirectory, managerRels[0].Attributes["Target"].Value);
            var partStream = zipReader.GetEntry(partPath);

            var partDoc = new XmlDocument();
            partDoc.Load(partStream);

            XmlNode e = partDoc.DocumentElement;
            
            b2xtranslator.Tools.Utils.replaceOutdatedNamespaces(ref e);
            
            return (XmlElement)e;
        }


    }

}
