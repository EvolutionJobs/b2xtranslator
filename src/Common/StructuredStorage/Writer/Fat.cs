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
    /// Class which represents the fat of a structured storage.
    /// Author: math
    /// </summary>
    class Fat : AbstractFat
    {
        List<UInt32> _diFatEntries = new List<UInt32>();

    
        // Number of sectors used by the fat.
        UInt32 _numFatSectors;
        internal UInt32 NumFatSectors
        {
            get { return _numFatSectors; }
        }


        // Number of sectors used by the difat.
        UInt32 _numDiFatSectors;
        internal UInt32 NumDiFatSectors
        {
            get { return _numDiFatSectors; }
        }


        // Start sector of the difat.
        UInt32 _diFatStartSector;
        internal UInt32 DiFatStartSector
        {
            get { return _diFatStartSector; }            
        }


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="context">the current context</param>
        internal Fat(StructuredStorageContext context)
            : base(context)
        {
        }


        /// <summary>
        /// Writes the difat entries to the fat
        /// </summary>
        /// <param name="sectorCount">Number of difat sectors.</param>
        /// <returns>Start sector of the difat.</returns>
        private UInt32 writeDiFatEntriesToFat(UInt32 sectorCount)
        {
            if (sectorCount == 0)
            {
                return SectorId.ENDOFCHAIN;
            }

            UInt32 startSector = _currentEntry;

            for (int i = 0; i < sectorCount; i++)
            {
                _currentEntry++;
                _entries.Add(SectorId.DIFSECT);
            }

            return startSector;
        }


        /// <summary>
        /// Writes the difat sectors to the output stream of the current context
        /// </summary>
        /// <param name="fatStartSector"></param>
        private void writeDiFatSectorsToStream(UInt32 fatStartSector)
        {
            // Add all entries of the difat
            for (UInt32 i = 0; i < _numFatSectors; i++)
            {
                _diFatEntries.Add(fatStartSector + i);
            }

            // Write the first 109 entries into the header
            for (int i = 0; i < 109; i++)
            {
                if (i < _diFatEntries.Count)
                {
                    _context.Header.writeNextDiFatSector(_diFatEntries[i]);
                }
                else
                {
                    _context.Header.writeNextDiFatSector(SectorId.FREESECT);
                }
            }

            if (_diFatEntries.Count <= 109)
            {
                return;
            }

            // handle remaining difat entries 

            List<UInt32> greaterDiFatEntries = new List<UInt32>();
            
            for (int i = 0; i < _diFatEntries.Count - 109; i++)
            {
                greaterDiFatEntries.Add(_diFatEntries[i + 109]);
            }          

            UInt32 diFatLink = _diFatStartSector + 1;
            int addressesInSector = _context.Header.SectorSize / 4;
            int sectorSplit = addressesInSector;

            // split difat at sector boundary and add link to next difat sector
            while (greaterDiFatEntries.Count >= sectorSplit)
            {
                greaterDiFatEntries.Insert(sectorSplit-1, diFatLink);
                diFatLink++;
                sectorSplit += addressesInSector;
            }

            // pad sector
            for (int i = greaterDiFatEntries.Count; i % (_context.Header.SectorSize / 4) != 0; i++)
            {
                greaterDiFatEntries.Add(SectorId.FREESECT);
            }
            greaterDiFatEntries.RemoveAt(greaterDiFatEntries.Count - 1);
            greaterDiFatEntries.Add(SectorId.ENDOFCHAIN);

            List<byte> output = _context.InternalBitConverter.getBytes(greaterDiFatEntries);

            // consistency check
            if (output.Count % _context.Header.SectorSize != 0)
            {
                throw new DiFatInconsistentException();
            }

            // write remaining difat sectors to stream
            _context.TempOutputStream.writeSectors(output.ToArray(), _context.Header.SectorSize, SectorId.FREESECT);

        }


        /// <summary>
        /// Marks the difat and fat sectors in the fat and writes the difat and fat data to the output stream of the current context.
        /// </summary>
        internal override void write()
        {

            // calculation of _numFatSectors and _numDiFatSectors (depending on each other)
            _numDiFatSectors = 0;            
            while (true)
            {
                UInt32 numDiFatSectorsOld = _numDiFatSectors;
                _numFatSectors = (UInt32)Math.Ceiling((double)(_entries.Count * 4) / (double)_context.Header.SectorSize) + _numDiFatSectors;
                _numDiFatSectors = (_numFatSectors <= 109) ? 0 : (UInt32)Math.Ceiling((double)((_numFatSectors - 109) * 4) / (double)(_context.Header.SectorSize - 1));
                if (numDiFatSectorsOld == _numDiFatSectors)
                {
                    break;
                }                
            }

            // writeDiFat
            _diFatStartSector = writeDiFatEntriesToFat(_numDiFatSectors);           
            writeDiFatSectorsToStream(_currentEntry);
           
            // Denote Fat entries in Fat
            for (int i = 0; i < _numFatSectors; i++)
            {
                _entries.Add(SectorId.FATSECT);
            }

            // write Fat
            _context.TempOutputStream.writeSectors((_context.InternalBitConverter.getBytes(_entries)).ToArray(), _context.Header.SectorSize, SectorId.FREESECT);
        }
    }
}
