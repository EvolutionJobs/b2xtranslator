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
using DIaLOGIKa.b2xtranslator.OpenXmlLib.DrawingML;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib.SpreadsheetML
{
    public class WorkbookPart : OpenXmlPart
    {
        private int _worksheetNumber;
        private int _chartsheetNumber;
        private int _drawingsNumber;
        private int _externalLinkNumber;
        protected WorksheetPart _workSheetPart;
        protected SharedStringPart _sharedStringPart;
        protected ExternalLinkPart _externalLinkPart;
        protected VbaProjectPart _vbaProjectPart;
        protected StylesPart _stylesPart;
        private string _type;

        public WorkbookPart(OpenXmlPartContainer parent, string contentType)
            : base(parent, 0)
        {
            this._worksheetNumber = 1;
            this._chartsheetNumber = 1;
            this._externalLinkNumber = 1;
            this._type = contentType;
        }

        public override string ContentType{
            get
            {
                return this._type;
            }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.OfficeDocument; }
        }

        /// <summary>
        /// returns the newly added worksheet part from the new excel document 
        /// </summary>
        /// <returns></returns>
        public WorksheetPart AddWorksheetPart()
        {
            this._workSheetPart = new WorksheetPart(this, this._worksheetNumber);
            this._worksheetNumber++;
            return this.AddPart(this._workSheetPart);
        }

        public ChartsheetPart AddChartsheetPart()
        {
            return this.AddPart(new ChartsheetPart(this, this._chartsheetNumber++));
        }

        public DrawingsPart AddDrawingsPart()
        {
            return this.AddPart(new DrawingsPart(this, this._drawingsNumber++));
        }

        /// <summary>
        /// returns the vba project part that contains the binary macro data
        /// </summary>
        public VbaProjectPart VbaProjectPart
        {
            get
            {
                if (this._vbaProjectPart == null)
                {
                    this._vbaProjectPart = this.AddPart(new VbaProjectPart(this));
                }
                return this._vbaProjectPart;
            }
        }

        /// <summary>
        /// return the latest created worksheetpart
        /// </summary>
        /// <returns></returns>
        public WorksheetPart GetWorksheetPart()
        {
            return this._workSheetPart; 
        }

        /// <summary>
        /// returns the worksheet part from the new excel document 
        /// </summary>
        /// <returns></returns>
        public ExternalLinkPart AddExternalLinkPart()
        {
            this._externalLinkPart = new ExternalLinkPart(this, this._externalLinkNumber);
            this._externalLinkNumber++;
            return this.AddPart(this._externalLinkPart);
        }

        /// <summary>
        /// return the latest created worksheetpart
        /// </summary>
        /// <returns></returns>
        public ExternalLinkPart GetExternalLinkPart()
        {
            return this._externalLinkPart;
        }

        public override string TargetName { get { return "workbook"; } }
        public override string TargetDirectory { get { return "xl"; } }


        /// <summary>
        /// returns the sharedstringtable part from the new excel document 
        /// </summary>
        /// <returns></returns>
        public SharedStringPart AddSharedStringPart()
        {
            this._sharedStringPart = new SharedStringPart(this);
            return this.AddPart(this._sharedStringPart);
        }

        /// <summary>
        /// returns the sharedstringtable part from the new excel document 
        /// </summary>
        /// <returns></returns>
        public StylesPart AddStylesPart()
        {
            this._stylesPart = new StylesPart(this);
            return this.AddPart(this._stylesPart);
        }

        internal int DrawingsNumber
        {
            get { return this._drawingsNumber; }
            set { this._drawingsNumber = value; }
        }

    }
}

