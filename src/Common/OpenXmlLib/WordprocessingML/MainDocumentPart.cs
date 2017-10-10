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

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML
{
    public class MainDocumentPart : ContentPart
    {
        protected StyleDefinitionsPart _styleDefinitionsPart;
        protected FontTablePart _fontTablePart;
        protected NumberingDefinitionsPart _numberingDefinitionsPart;
        protected SettingsPart _settingsPart;
        protected FootnotesPart _footnotesPart;
        protected EndnotesPart _endnotesPart;
        protected CommentsPart _commentsPart;
        protected VbaProjectPart _vbaProjectPart;
        protected GlossaryPart _glossaryPart;
        protected KeyMapCustomizationsPart _customizationsPart;

        protected int _headerPartCount = 0;
        protected int _footerPartCount = 0;

        private string _contentType = WordprocessingMLContentTypes.MainDocument;
        
        public MainDocumentPart(OpenXmlPartContainer parent, string contentType)
            : base(parent)
        {
            this._contentType = contentType;
        }

        public override string ContentType { get { return this._contentType; } }
        public override string RelationshipType { get { return OpenXmlRelationshipTypes.OfficeDocument; } }
        public override string TargetName { get { return "document"; } }
        public override string TargetDirectory { get { return "word"; } }

        // unique parts

        public KeyMapCustomizationsPart CustomizationsPart
        {
            get
            {
                if (this._customizationsPart == null)
                {
                    this._customizationsPart = new KeyMapCustomizationsPart(this);
                    this.AddPart(this._customizationsPart);
                }
                return this._customizationsPart;
            }
        }

        public GlossaryPart GlossaryPart
        {
            get
            {
                if (this._glossaryPart == null)
                {
                    this._glossaryPart = new GlossaryPart(this, WordprocessingMLContentTypes.Glossary);
                    this.AddPart(this._glossaryPart);
                }
                return this._glossaryPart;
            }
        }

        public StyleDefinitionsPart StyleDefinitionsPart
        {
            get
            {
                if (this._styleDefinitionsPart == null)
                {
                    this._styleDefinitionsPart = new StyleDefinitionsPart(this);
                    this.AddPart(this._styleDefinitionsPart);
                }
                return this._styleDefinitionsPart;
            }
        }

        public SettingsPart SettingsPart
        {
            get
            {
                if (this._settingsPart == null)
                {
                    this._settingsPart = new SettingsPart(this);
                    this.AddPart(this._settingsPart);
                }
                return this._settingsPart;
            }
        }

        public NumberingDefinitionsPart NumberingDefinitionsPart
        {
            get
            {
                if (this._numberingDefinitionsPart == null)
                {
                    this._numberingDefinitionsPart = new NumberingDefinitionsPart(this);
                    this.AddPart(this._numberingDefinitionsPart);
                }
                return this._numberingDefinitionsPart;
            }
        }

        public FontTablePart FontTablePart
        {
            get
            {
                if (this._fontTablePart == null)
                {
                    this._fontTablePart = new FontTablePart(this);
                    this.AddPart(this._fontTablePart);
                }
                return this._fontTablePart;
            }
        }

        public EndnotesPart EndnotesPart
        {
            get
            {
                if (this._endnotesPart == null)
                {
                    this._endnotesPart = new EndnotesPart(this);
                    this.AddPart(this._endnotesPart);
                }
                return this._endnotesPart;
            }
        }

        public FootnotesPart FootnotesPart
        {
            get
            {
                if (this._footnotesPart == null)
                {
                    this._footnotesPart = new FootnotesPart(this);
                    this.AddPart(this._footnotesPart);
                }
                return this._footnotesPart;
            }
        }

        public CommentsPart CommentsPart
        {
            get 
            {
                if (this._commentsPart == null)
                {
                    this._commentsPart = new CommentsPart(this);
                    this.AddPart(this._commentsPart);
                }
                return this._commentsPart;
            }
        }

        public VbaProjectPart VbaProjectPart
        {
            get 
            {
                if(this._vbaProjectPart == null)
                {
                    this._vbaProjectPart = this.AddPart(new VbaProjectPart(this));
                }
                return this._vbaProjectPart;
            }
        }

        // non unique parts

        public HeaderPart AddHeaderPart()
        {
            return this.AddPart(new HeaderPart(this, ++this._headerPartCount));
        }

        public FooterPart AddFooterPart()
        {
            return this.AddPart(new FooterPart(this, ++this._footerPartCount));
        }
    }
}
