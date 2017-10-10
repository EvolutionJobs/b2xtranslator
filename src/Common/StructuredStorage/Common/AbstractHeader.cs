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

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Common
{
    /// <summary>
    /// Abstract class fo the header of a compound file.
    /// Author: math
    /// </summary>
    abstract internal class AbstractHeader
    {
        protected const UInt64 MAGIC_NUMBER = 0xE11AB1A1E011CFD0;

        protected AbstractIOHandler _ioHandler;

        // Sector shift and sector size
        private ushort _sectorShift;
        public ushort SectorShift
        {
            get { return _sectorShift; }
            set
            {
                _sectorShift = value;
                // Calculate sector size
                _sectorSize = (ushort)Math.Pow((double)2, (double)_sectorShift);
                if (_sectorShift != 9 && _sectorShift != 12)
                {
                    throw new UnsupportedSizeException("SectorShift: " + _sectorShift);
                }
            }
        }
        private ushort _sectorSize;
        public ushort SectorSize
        {
            get { return _sectorSize; }
        }


        // Minisector shift and Minisector size
        private ushort _miniSectorShift;
        public ushort MiniSectorShift
        {
            get { return _miniSectorShift; }
            set
            {
                _miniSectorShift = value;
                // Calculate mini sector size
                _miniSectorSize = (ushort)Math.Pow((double)2, (double)_miniSectorShift);
                if (_miniSectorShift != 6)
                {
                    throw new UnsupportedSizeException("MiniSectorShift: " + _miniSectorShift);
                }
            }
        }
        private ushort _miniSectorSize;
        public ushort MiniSectorSize
        {
            get { return _miniSectorSize; }
        }


        // CSectDir
        private uint _noSectorsInDirectoryChain4KB;
        public uint NoSectorsInDirectoryChain4KB
        {
            get { return _noSectorsInDirectoryChain4KB; }
            set
            {
                if (_sectorSize == 512 && value != 0)
                {
                    throw new ValueNotZeroException("_csectDir");
                }
                _noSectorsInDirectoryChain4KB = value;
            }
        }


        // CSectFat
        private uint _noSectorsInFatChain;
        public uint NoSectorsInFatChain
        {
            get { return _noSectorsInFatChain; }
            set
            {
                _noSectorsInFatChain = value;
                if (value > _ioHandler.IOStreamSize / SectorSize)
                {
                    throw new InvalidValueInHeaderException("NoSectorsInFatChain");
                }

            }
        }


        // SectDirStart
        private uint _directoryStartSector;
        public uint DirectoryStartSector
        {
            get { return _directoryStartSector; }
            set
            {
                _directoryStartSector = value;
                if (value > _ioHandler.IOStreamSize / SectorSize && value != SectorId.ENDOFCHAIN)
                {
                    throw new InvalidValueInHeaderException("DirectoryStartSector");
                }
            }
        }


        // UInt32ULMiniSectorCutoff
        private uint _miniSectorCutoff;
        public uint MiniSectorCutoff
        {
            get { return _miniSectorCutoff; }
            set
            {
                _miniSectorCutoff = value;
                if (value != 0x1000)
                {
                    throw new UnsupportedSizeException("MiniSectorCutoff");
                }
            }
        }



        // SectMiniFatStart
        private uint _miniFatStartSector;
        public uint MiniFatStartSector
        {
            get { return _miniFatStartSector; }
            set
            {
                _miniFatStartSector = value;
                if (value > _ioHandler.IOStreamSize / SectorSize && value != SectorId.ENDOFCHAIN)
                {
                    throw new InvalidValueInHeaderException("MiniFatStartSector");
                }
            }
        }


        // CSectMiniFat
        private uint _noSectorsInMiniFatChain;
        public uint NoSectorsInMiniFatChain
        {
            get { return _noSectorsInMiniFatChain; }
            set
            {
                _noSectorsInMiniFatChain = value;
                if (value > _ioHandler.IOStreamSize / SectorSize)
                {
                    throw new InvalidValueInHeaderException("NoSectorsInMiniFatChain");
                }

            }
        }


        // SectDifStart
        private uint _diFatStartSector;
        public uint DiFatStartSector
        {
            get { return _diFatStartSector; }
            set
            {
                _diFatStartSector = value;
                if (value > _ioHandler.IOStreamSize / SectorSize && value != SectorId.ENDOFCHAIN && value != SectorId.FREESECT)
                {
                    throw new InvalidValueInHeaderException("DiFatStartSector", string.Format("Details: value={0};_ioHandler.IOStreamSize={1};SectorSize={2}; SectorId.ENDOFCHAIN: {3}", value, _ioHandler.IOStreamSize, SectorSize, SectorId.ENDOFCHAIN));
                }
            }
        }


        // CSectDif
        private uint _noSectorsInDiFatChain;
        public uint NoSectorsInDiFatChain
        {
            get { return _noSectorsInDiFatChain; }
            set
            {
                _noSectorsInDiFatChain = value;
                if (value > _ioHandler.IOStreamSize / SectorSize)
                {
                    throw new InvalidValueInHeaderException("NoSectorsInDiFatChain");
                }

            }
        }

    }
}
