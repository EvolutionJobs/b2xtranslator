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
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Reader
{
    /// <summary>
    /// Represents the Fat in a compound file
    /// Author: math
    /// </summary>
    internal class Fat : AbstractFat
    {        
        List<UInt32> _sectorsUsedByFat = new List<UInt32>();
        List<UInt32> _sectorsUsedByDiFat = new List<UInt32>();        

        override internal UInt16 SectorSize
        {
            get { return _header.SectorSize; }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="header">Handle to the header of the compound file</param>
        /// <param name="fileHandler">Handle to the file handler of the compound file</param>
        internal Fat(Header header, InputHandler fileHandler)
            : base(header, fileHandler)
        {            
            Init();
        }


        /// <summary>
        /// Seeks to a given position in a sector
        /// </summary>
        /// <param name="sector">The sector to seek to</param>
        /// <param name="position">The position in the sector to seek to</param>
        /// <returns>The new position in the stream.</returns>
        override internal long SeekToPositionInSector(long sector, long position)
        {
            return _fileHandler.SeekToPositionInSector(sector, position);
        }


        /// <summary>
        /// Returns the next sector in a chain
        /// </summary>
        /// <param name="currentSector">The current sector in the chain</param>
        /// <returns>The next sector in the chain</returns>
        override protected UInt32 GetNextSectorInChain(UInt32 currentSector)
        {
            UInt32 sectorInFile = _sectorsUsedByFat[(int)(currentSector / _addressesPerSector)];
            // calculation of position:
            // currentSector % _addressesPerSector = number of address in the sector address
            // address uses 32 bit = 4 bytes
            _fileHandler.SeekToPositionInSector(sectorInFile, 4 * (currentSector % _addressesPerSector));
            return _fileHandler.ReadUInt32();
        }
        

        /// <summary>
        /// Initalizes the Fat
        /// </summary>
        private void Init()
        {
            ReadFirst109SectorsUsedByFAT();
            ReadSectorsUsedByFatFromDiFat();            
            CheckConsistency();
        }


        /// <summary>
        /// Reads the first 109 sectors of the Fat stored in the header
        /// </summary>
        private void ReadFirst109SectorsUsedByFAT()
        {
            // Header sector: -1
            _fileHandler.SeekToPositionInSector(-1, 0x4C);
            UInt32 fatSector;
            for (int i = 0; i < 109; i++)
            {
                fatSector = _fileHandler.ReadUInt32();
                if (fatSector == SectorId.FREESECT)
                {
                    break;
                }
                _sectorsUsedByFat.Add(fatSector);
            }
        }


        /// <summary>
        /// Reads the sectors of the Fat which are stored in the DiFat
        /// </summary>
        private void ReadSectorsUsedByFatFromDiFat()
        {
            if (_header.DiFatStartSector == SectorId.ENDOFCHAIN || _header.NoSectorsInDiFatChain == 0x0)
            {
                return;
            }

            _fileHandler.SeekToSector(_header.DiFatStartSector);
            bool lastFatSectorFound = false;
            _sectorsUsedByDiFat.Add(_header.DiFatStartSector);

            while (true)
            {
                // Add all addresses contained in the current difat sector except the last address (it points to next difat sector)
                for (int i = 0; i < _addressesPerSector - 1; i++)
                {
                    UInt32 fatSector = _fileHandler.ReadUInt32();
                    if (fatSector == SectorId.FREESECT)
                    {
                        lastFatSectorFound = true;
                        break;
                    }
                    _sectorsUsedByFat.Add(fatSector);
                }

                if (lastFatSectorFound)
                {
                    break;
                }

                // Last address in difat sector points to next difat sector
                UInt32 nextDiFatSector = _fileHandler.ReadUInt32();
                if (nextDiFatSector == SectorId.FREESECT || nextDiFatSector == SectorId.ENDOFCHAIN)
                {
                    break;
                }

                _sectorsUsedByDiFat.Add(nextDiFatSector);
                _fileHandler.SeekToSector(nextDiFatSector);

                if (_sectorsUsedByDiFat.Count > _header.NoSectorsInDiFatChain)
                {
                    throw new ChainSizeMismatchException("DiFat");
                }
            }
        }


        /// <summary>
        /// Checks whether the sizes specified in the header matches the actual sizes
        /// </summary>
        private void CheckConsistency()
        {
            if (_sectorsUsedByDiFat.Count != _header.NoSectorsInDiFatChain
                || _sectorsUsedByFat.Count != _header.NoSectorsInFatChain)
            {
                throw new ChainSizeMismatchException("Fat/DiFat");
            }
        }
    }
}
