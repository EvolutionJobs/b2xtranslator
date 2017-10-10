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
using System.IO;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Reader
{
    /// <summary>
    /// Abstract class of a Fat in a compound file
    /// Author: math
    /// </summary>
    internal abstract class AbstractFat
    {
        protected Header _header;
        protected InputHandler _fileHandler;
        protected int _addressesPerSector;



        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="header">Handle to the header of the compound file</param>
        /// <param name="fileHandler">Handle to the file handler of the compound file</param>
        internal AbstractFat(Header header, InputHandler fileHandler)
        {
            _header = header;
            _fileHandler = fileHandler;
            _addressesPerSector = (int)_header.SectorSize / 4;
        }

        internal Stream _InternalFileStream
        {
            get
            {
                return _fileHandler._InternalFileStream;
            }
        }


        /// <summary>
        /// Returns the sectors in a chain which starts at a given sector
        /// </summary>
        /// <param name="startSector">The start sector of the chain</param>
        /// <param name="maxCount">The maximum count of sectors in a chain</param>
        /// <param name="name">The name of a chain</param>
        internal List<UInt32> GetSectorChain(UInt32 startSector, UInt64 maxCount, string name)
        {
            return GetSectorChain(startSector, maxCount, name, false);
        }

        /// <summary>
        /// Returns the sectors in a chain which starts at a given sector
        /// </summary>
        /// <param name="startSector">The start sector of the chain</param>
        /// <param name="maxCount">The maximum count of sectors in a chain</param>
        /// <param name="name">The name of a chain</param>
        /// <param name="immediateCycleCheck">Flag whether to check for cycles in every loop</param>
        internal List<UInt32> GetSectorChain(UInt32 startSector, UInt64 maxCount, string name, bool immediateCycleCheck)
        {
            var result = new List<UInt32>();

            result.Add(startSector);
            while (true)
            {
                var nextSectorInStream = this.GetNextSectorInChain(result[result.Count - 1]);

                // Check for invalid sectors in chain
                if (nextSectorInStream == SectorId.DIFSECT || nextSectorInStream == SectorId.FATSECT || nextSectorInStream == SectorId.FREESECT)
                {
                    throw new InvalidSectorInChainException();
                }
                if (nextSectorInStream == SectorId.ENDOFCHAIN)
                {
                    break;
                }
                                
                if (immediateCycleCheck)
                {
                    if (result.Contains(nextSectorInStream))
                    {
                        throw new ChainCycleDetectedException(name);
                    }
                }

                result.Add(nextSectorInStream);
                
                // Chain too long
                if ((UInt64)(result.Count) > maxCount)
                {
                    throw new ChainSizeMismatchException(name);
                }
            }

            return result;
        }


        /// <summary>
        /// Reads bytes into an array
        /// </summary>
        /// <param name="array">The array to read to</param>
        /// <param name="offset">The offset in the array to read to</param>
        /// <param name="count">The number of bytes to read</param>
        /// <returns>The number of bytes read</returns>
        internal int UncheckedRead(byte[] array, int offset, int count)
        {
            return _fileHandler.UncheckedRead(array, offset, count);
        }

        /// <summary>
        /// Reads a byte at the current position of the file stream.
        /// Advances the stream pointer accordingly.
        /// </summary>
        internal int UncheckedReadByte()
        {
            return _fileHandler.UncheckedReadByte();
        }


        /// <summary>
        /// Returns the next sector in a chain
        /// </summary>
        /// <param name="currentSector">The current sector in the chain</param>
        /// <returns>The next sector in the chain</returns>
        abstract protected UInt32 GetNextSectorInChain(UInt32 currentSector);


        /// <summary>
        /// Seeks to a given position in a sector
        /// </summary>
        /// <param name="sector">The sector to seek to</param>
        /// <param name="position">The position in the sector to seek to</param>
        /// <returns></returns>
        abstract internal long SeekToPositionInSector(long sector, long position);
        abstract internal ushort SectorSize { get; }
    }
}
