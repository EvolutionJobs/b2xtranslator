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
    public class SSTData: IVisitable
    {
        /// <summary>
        /// Total and unique number of strings in this SST-Biffrecord 
        /// </summary>
        public uint cstTotal;
        public uint cstUnique;

        /// <summary>
        /// Two lists to store the shared String Data 
        /// </summary>
        public List<string> StringList;
        public List<StringFormatAssignment> FormatList;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="sst">The SST BiffRecord</param>
        public SSTData(SST sst)
        {
            this.copySSTData(sst); 
        }

        /// <summary>
        /// copies the different datasources from the SST BiffRecord 
        /// </summary>
        /// <param name="sst">The SST BiffRecord </param>
        public void copySSTData(SST sst)
        {
            this.StringList = sst.StringList;
            this.FormatList = sst.FormatList;
            this.cstTotal = sst.cstTotal;
            this.cstUnique = sst.cstUnique; 
        }

        public List<StringFormatAssignment> getFormatingRuns(int stringNumber)
        {
            var returnList = new List<StringFormatAssignment>();
            foreach (var item in this.FormatList)
            {
                if (item.StringNumber == stringNumber)
                {
                    returnList.Add(item); 
                }
                
            }
            return returnList; 
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<SSTData>)mapping).Apply(this);
        }

        #endregion
    }
}
