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
            _contentType = contentType;
        }

        public override string ContentType { get { return _contentType; } }
        public override string RelationshipType { get { return OpenXmlRelationshipTypes.OfficeDocument; } }
        public override string TargetName { get { return "document"; } }
        public override string TargetDirectory { get { return "word"; } }

        // unique parts

        public KeyMapCustomizationsPart CustomizationsPart
        {
            get
            {
                if (_customizationsPart == null)
                {
                    _customizationsPart = new KeyMapCustomizationsPart(this);
                    this.AddPart(_customizationsPart);
                }
                return _customizationsPart;
            }
        }

        public GlossaryPart GlossaryPart
        {
            get
            {
                if (_glossaryPart == null)
                {
                    _glossaryPart = new GlossaryPart(this, WordprocessingMLContentTypes.Glossary);
                    this.AddPart(_glossaryPart);
                }
                return _glossaryPart;
            }
        }

        public StyleDefinitionsPart StyleDefinitionsPart
        {
            get
            {
                if (_styleDefinitionsPart == null)
                {
                    _styleDefinitionsPart = new StyleDefinitionsPart(this);
                    this.AddPart(_styleDefinitionsPart);
                }
                return _styleDefinitionsPart;
            }
        }

        public SettingsPart SettingsPart
        {
            get
            {
                if (_settingsPart == null)
                {
                    _settingsPart = new SettingsPart(this);
                    this.AddPart(_settingsPart);
                }
                return _settingsPart;
            }
        }

        public NumberingDefinitionsPart NumberingDefinitionsPart
        {
            get
            {
                if (_numberingDefinitionsPart == null)
                {
                    _numberingDefinitionsPart = new NumberingDefinitionsPart(this);
                    this.AddPart(_numberingDefinitionsPart);
                }
                return _numberingDefinitionsPart;
            }
        }

        public FontTablePart FontTablePart
        {
            get
            {
                if (_fontTablePart == null)
                {
                    _fontTablePart = new FontTablePart(this);
                    this.AddPart(_fontTablePart);
                }
                return _fontTablePart;
            }
        }

        public EndnotesPart EndnotesPart
        {
            get
            {
                if (_endnotesPart == null)
                {
                    _endnotesPart = new EndnotesPart(this);
                    this.AddPart(_endnotesPart);
                }
                return _endnotesPart;
            }
        }

        public FootnotesPart FootnotesPart
        {
            get
            {
                if (_footnotesPart == null)
                {
                    _footnotesPart = new FootnotesPart(this);
                    this.AddPart(_footnotesPart);
                }
                return _footnotesPart;
            }
        }

        public CommentsPart CommentsPart
        {
            get 
            {
                if (_commentsPart == null)
                {
                    _commentsPart = new CommentsPart(this);
                    this.AddPart(_commentsPart);
                }
                return _commentsPart;
            }
        }

        public VbaProjectPart VbaProjectPart
        {
            get 
            {
                if(_vbaProjectPart == null)
                {
                    _vbaProjectPart = this.AddPart(new VbaProjectPart(this));
                }
                return _vbaProjectPart;
            }
        }

        // non unique parts

        public HeaderPart AddHeaderPart()
        {
            return this.AddPart(new HeaderPart(this, ++_headerPartCount));
        }

        public FooterPart AddFooterPart()
        {
            return this.AddPart(new FooterPart(this, ++_footerPartCount));
        }
    }
}
