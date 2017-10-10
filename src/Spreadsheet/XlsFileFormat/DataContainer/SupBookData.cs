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

using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;


namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class contains several information about the SUPBOOK BIFF Record 
    /// 
    /// </summary>
    public class SupBookData : IVisitable
    {
        private string virtPath;
        public string VirtPath
        {
            get { return this.virtPath; }
        }

        private string[] rgst;
        public string[] RGST
        {
            get { return this.rgst; }
        }

        private bool selfref;

        public bool SelfRef
        {
            get { return this.selfref; }
        }

        private LinkedList<XCTData> xctDataList;
        public LinkedList<XCTData> XCTDataList
        {
            get { return this.xctDataList; }
        }

        private LinkedList<string> externNames;
        public LinkedList<string> ExternNames
        {
            get { return this.externNames; }

        }

        public int ExternalLinkId;
        public string ExternalLinkRef;
        public int Number; 

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="supbook">SUPBOOK BIFF Record </param>
        public SupBookData(SupBook supbook)
        {
            this.rgst = supbook.rgst;
            this.virtPath = supbook.virtpathstring;
            this.selfref = supbook.isselfreferencing;
            this.xctDataList = new LinkedList<XCTData>();
            this.externNames = new LinkedList<string>(); 
        }

        /// <summary>
        /// returns the value at the specified position
        /// </summary>
        /// <param name="index">searched index</param>
        /// <returns></returns>
        public string getRgstString(int index)
        {
            return this.rgst[index]; 
        }

        /// <summary>
        /// Add a XCT Data structure to the internal stack 
        /// </summary>
        /// <param name="xct"></param>
        public void addXCT(XCT xct)
        {
            var xctdata = new XCTData(xct);
            this.xctDataList.AddLast(xctdata); 
        }

        public void addCRN(CRN crn)
        {
            this.xctDataList.Last.Value.addCRN(crn);           
        }

        public void addEXTERNNAME(ExternName extname)
        {
            this.externNames.AddLast(extname.extName); 
        }


        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<SupBookData>)mapping).Apply(this);
        }

        #endregion


    }
}
