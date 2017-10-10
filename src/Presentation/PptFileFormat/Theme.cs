/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
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
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Xml;
using DIaLOGIKa.b2xtranslator.ZipUtils;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
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
        protected override XmlElement ExtractDocumentElement(ZipReader zipReader, XmlNodeList rootRels)
        {
            if (rootRels.Count != 1)
                throw new Exception("Expected actly one Relationship in Theme OOXML doc");

            var managerPath = rootRels[0].Attributes["Target"].Value;
            var managerDirectory = Path.GetDirectoryName(managerPath).Replace("\\", "/");
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

            var partPath = string.Format("{0}/{1}", managerDirectory, managerRels[0].Attributes["Target"].Value);
            var partStream = zipReader.GetEntry(partPath);

            var partDoc = new XmlDocument();
            partDoc.Load(partStream);

            XmlNode e = partDoc.DocumentElement;
            
            DIaLOGIKa.b2xtranslator.Tools.Utils.replaceOutdatedNamespaces(ref e);
            
            return (XmlElement)e;
        }


    }

}
