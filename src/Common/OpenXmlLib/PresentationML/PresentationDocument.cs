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
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML
{
    public class PresentationDocument : OpenXmlPackage
    {
        protected PresentationPart _presentationPart;
        protected OpenXmlPackage.DocumentType _documentType;

        protected PresentationDocument(string fileName, OpenXmlPackage.DocumentType type)
            : base(fileName)
        {
            switch (type)
            {
                case OpenXmlPackage.DocumentType.Document:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.Presentation);
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledDocument:
                    this._presentationPart = new PresentationPart(this, PresentationMLContentTypes.PresentationMacro);
                    break;
                case OpenXmlPackage.DocumentType.Template:
                    break;
                case OpenXmlPackage.DocumentType.MacroEnabledTemplate:
                    break;
            }

            this.AddPart(_presentationPart);
        }

        public static PresentationDocument Create(string fileName, OpenXmlPackage.DocumentType type)
        {
            PresentationDocument presentation = new PresentationDocument(fileName, type);

            return presentation;
        }

        public PresentationPart PresentationPart
        {
            get { return _presentationPart; }
        }
    }
}
