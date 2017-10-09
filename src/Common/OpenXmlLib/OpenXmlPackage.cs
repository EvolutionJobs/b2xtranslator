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

using DIaLOGIKa.b2xtranslator.ZipUtils;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.IO;
using System.Collections.Generic;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib
{
    public abstract class OpenXmlPackage : OpenXmlPartContainer, IDisposable
    {
        #region Protected members
        protected string _fileName;

        protected Dictionary<string, string> _defaultTypes = new Dictionary<string, string>();
        protected Dictionary<string, string> _partOverrides = new Dictionary<string, string>();

        protected CorePropertiesPart _coreFilePropertiesPart;
        protected AppPropertiesPart _appPropertiesPart;

        protected int _imageCounter;
        protected int _vmlCounter;
        protected int _oleCounter;
        #endregion

        public enum DocumentType
        {
            Document,
            MacroEnabledDocument,
            MacroEnabledTemplate,
            Template
        }

        protected OpenXmlPackage(string fileName)
        {
            _fileName = fileName;

            _defaultTypes.Add("rels", OpenXmlContentTypes.Relationships);
            _defaultTypes.Add("xml", OpenXmlContentTypes.Xml);
            _defaultTypes.Add("bin", OpenXmlContentTypes.OleObject);
            _defaultTypes.Add("vml", OpenXmlContentTypes.Vml);
            _defaultTypes.Add("emf", OpenXmlContentTypes.Emf);
            _defaultTypes.Add("wmf", OpenXmlContentTypes.Wmf);
        }

        public void Dispose()
        {
            this.Close();
        }


        public virtual void Close()
        {
            // serialize the package on closing
            OpenXmlWriter writer = new OpenXmlWriter();
            writer.Open(this.FileName);

            this.WritePackage(writer);

            writer.Close();
        }

        public string FileName
        {
            get { return _fileName; }
            set { _fileName = value; }
        }

        public CorePropertiesPart CoreFilePropertiesPart
        {
            get { return _coreFilePropertiesPart; }
            set { _coreFilePropertiesPart = value; }
        }

        public CorePropertiesPart AddCoreFilePropertiesPart()
        {
            this.CoreFilePropertiesPart = new CorePropertiesPart(this);
            return this.AddPart(this.CoreFilePropertiesPart);
        }

        public AppPropertiesPart AppPropertiesPart
        {
            get { return _appPropertiesPart; }
            set { _appPropertiesPart = value; }
        }
               
        public AppPropertiesPart AddAppPropertiesPart()
        {
            this.AppPropertiesPart = new AppPropertiesPart(this);
            return this.AddPart(this.AppPropertiesPart);
        }

        internal void AddContentTypeDefault(string extension, string contentType)
        {
            if (!_defaultTypes.ContainsKey(extension))
                _defaultTypes.Add(extension, contentType);
        }

        internal void AddContentTypeOverride(string partNameAbsolute, string contentType)
        {
            if (!_partOverrides.ContainsKey(partNameAbsolute))
                _partOverrides.Add(partNameAbsolute, contentType);
        }

        internal int GetNextImageId()
        {
            _imageCounter++;
            return _imageCounter;
        }

        internal int GetNextVmlId()
        {
            _vmlCounter++;
            return _vmlCounter;
        }

        internal int GetNextOleId()
        {
            _oleCounter++;
            return _oleCounter;
        }

        protected void WritePackage(OpenXmlWriter writer)
        {
            foreach (OpenXmlPart part in this.Parts)
            {
                part.WritePart(writer);
            }

            this.WriteRelationshipPart(writer);

            // write content types
            writer.AddPart("[Content_Types].xml");

            writer.WriteStartDocument();
            writer.WriteStartElement("Types", OpenXmlNamespaces.ContentTypes);

            foreach (string extension in _defaultTypes.Keys)
            {
                writer.WriteStartElement("Default", OpenXmlNamespaces.ContentTypes);
                writer.WriteAttributeString("Extension", extension);
                writer.WriteAttributeString("ContentType", _defaultTypes[extension]);
                writer.WriteEndElement();
            }

            foreach (string partName in _partOverrides.Keys)
            {
                writer.WriteStartElement("Override", OpenXmlNamespaces.ContentTypes);
                writer.WriteAttributeString("PartName", partName);
                writer.WriteAttributeString("ContentType", _partOverrides[partName]);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndDocument();
        }
    }
}
