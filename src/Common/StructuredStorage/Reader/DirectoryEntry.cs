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
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Reader
{
    /// <summary>
    /// Encapsulates a directory entry
    /// Author: math
    /// </summary>
    public class DirectoryEntry : AbstractDirectoryEntry
    {
        InputHandler _fileHandler;
        Header _header;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="header">Handle to the header of the compound file</param>
        /// <param name="fileHandler">Handle to the file handler of the compound file</param>
        /// <param name="sid">The sid of the directory entry</param>
        internal DirectoryEntry(Header header, InputHandler fileHandler, UInt32 sid, string path) : base(sid)
        {
            _header = header;
            _fileHandler = fileHandler;
            //_sid = sid;            
            ReadDirectoryEntry();
            _path = path;
        }


        /// <summary>
        /// Reads the values of the directory entry. The position of the file handler must be at the start of a directory entry.
        /// </summary>
        private void ReadDirectoryEntry()
        {     
            Name = _fileHandler.ReadString(64);

            // Name length check: lengthOfName = length of the element in bytes including Unicode NULL
            UInt16 lengthOfName = _fileHandler.ReadUInt16();
            // Commented out due to trouble with odd unicode-named streams in PowerPoint -- flgr
            /*if (lengthOfName != (_name.Length + 1) * 2)
            {
                throw new InvalidValueInDirectoryEntryException("_cb");
            }*/
            // Added warning - math
            if (lengthOfName != (_name.Length + 1) * 2)
            {
                TraceLogger.Warning("Length of the name (_cb) of stream '" + Name + "' is not correct.");
            }


            Type = (DirectoryEntryType)_fileHandler.ReadByte();
            Color = (DirectoryEntryColor)_fileHandler.ReadByte();
            LeftSiblingSid = _fileHandler.ReadUInt32();
            RightSiblingSid = _fileHandler.ReadUInt32();
            ChildSiblingSid = _fileHandler.ReadUInt32();

            byte[] array = new byte[16];
            _fileHandler.Read(array);
            ClsId = new Guid(array);

            UserFlags = _fileHandler.ReadUInt32();
            // Omit creation time
            _fileHandler.ReadUInt64();
            // Omit modification time 
            _fileHandler.ReadUInt64();
            StartSector = _fileHandler.ReadUInt32();

            UInt32 sizeLow = _fileHandler.ReadUInt32();
            UInt32 sizeHigh = _fileHandler.ReadUInt32();

            if (_header.SectorSize == 512 && sizeHigh != 0x0)
            {
                // Must be zero according to the specification. However, this requirement can be ommited.
                TraceLogger.Warning("ul_SizeHigh of stream '" + Name + "' should be zero as sector size is 512.");
                sizeHigh = 0x0;
            }
            SizeOfStream = ((UInt64)sizeHigh << 32) + sizeLow;
        }
    }
}
