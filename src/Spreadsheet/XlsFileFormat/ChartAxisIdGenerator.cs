/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using System.Collections.Generic;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// An internal helper class for generating unique axis ids to be used in OpenXML
    /// </summary>
    class ChartAxisIdGenerator
    {
        private int _id;

        /// <summary>
        /// A list containing all axis ids belonging to a chart group.
        /// </summary>
        private List<int> _idList = new List<int>();


        /// <summary>
        /// This class is a singleton
        /// </summary>
        private static ChartAxisIdGenerator _instance;

        private ChartAxisIdGenerator()
        {
        }

        public static ChartAxisIdGenerator Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ChartAxisIdGenerator();
                }
                return _instance;
            }
        }

        public void StartNewChartsheetSubstream()
        {
            _id = 0;
            _idList.Clear();
        }


        public void StartNewAxisGroup()
        {
            _idList.Clear();
        }

        public int GenerateId()
        {
            int newId = _id++;
            _idList.Add(newId);
            return newId;
        }

        public int[] AxisIds
        {
            get
            {
                var retVal = new int[_idList.Count];
                _idList.CopyTo(retVal);
                return retVal;
            }
        }
    }
}
