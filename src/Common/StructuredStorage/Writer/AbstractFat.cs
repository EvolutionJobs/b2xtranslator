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
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{
    /// <summary>
    /// Abstract class of a Fat in a compound file
    /// Author: math
    /// </summary>
    abstract class AbstractFat
    {
        protected List<uint> _entries = new List<uint>();
        protected uint _currentEntry = 0;
        protected StructuredStorageContext _context;


        /// <summary>
        /// Constructor
        /// <param name="context">the current context</param>
        /// </summary>
        protected AbstractFat(StructuredStorageContext context)
        {
            _context = context;
        }


        /// <summary>
        /// Write a chain to the fat.
        /// </summary>
        /// <param name="entryCount">number of entries in the chain</param>
        /// <returns></returns>
        internal uint writeChain(uint entryCount)
        {
            if (entryCount == 0)
            {
                return SectorId.FREESECT;
            }

            var startSector = _currentEntry;

            for (int i = 0; i < entryCount - 1; i++)
            {
                _currentEntry++;
                _entries.Add(_currentEntry);
            }

            _currentEntry++;
            _entries.Add(SectorId.ENDOFCHAIN);

            return startSector;
        }


        abstract internal void write();

    }
}
