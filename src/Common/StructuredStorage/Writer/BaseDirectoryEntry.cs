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
    /// Common base class for stream and storage directory entries
    /// Author: math
    /// </summary>
    abstract public class BaseDirectoryEntry : AbstractDirectoryEntry
    {
        private StructuredStorageContext _context;
        internal StructuredStorageContext Context
        {
            get { return _context; }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="name">Name of the directory entry.</param>
        /// <param name="context">the current context</param>
        internal BaseDirectoryEntry(string name, StructuredStorageContext context)            
        {
            _context = context;
            Name = name;
            setInitialValues();
        }


        /// <summary>
        /// Set the initial values
        /// </summary>
        private void setInitialValues()
        {
            this.ChildSiblingSid = SectorId.FREESECT;
            this.LeftSiblingSid = SectorId.FREESECT;
            this.RightSiblingSid = SectorId.FREESECT;
            this.ClsId = Guid.Empty;
            this.Color = DirectoryEntryColor.DE_BLACK;
            this.StartSector = 0x0;
            this.ClsId = Guid.Empty;
            this.UserFlags = 0x0;
            this.SizeOfStream = 0x0;
        }


        /// <summary>
        /// Writes the directory entry to the directory stream of the current context
        /// </summary>
        internal void write()
        {
            OutputHandler directoryStream = _context.DirectoryStream;
            char[] unicodeName = _name.ToCharArray();
            int paddingCounter = 0;
            foreach (UInt16 unicodeChar in  unicodeName)
            {
                directoryStream.writeUInt16(unicodeChar);
                paddingCounter++;
            }
            while (paddingCounter < 32)
            {
                directoryStream.writeUInt16(0x0);
                paddingCounter++;
            }
            directoryStream.writeUInt16(this.LengthOfName);
            directoryStream.writeByte((byte)this.Type);
            directoryStream.writeByte((byte)this.Color);
            directoryStream.writeUInt32(this.LeftSiblingSid);
            directoryStream.writeUInt32(this.RightSiblingSid);
            directoryStream.writeUInt32(this.ChildSiblingSid);
            directoryStream.write(this.ClsId.ToByteArray());
            directoryStream.writeUInt32(this.UserFlags);
            //FILETIME set to 0x0
            directoryStream.write(new byte[16]);

            directoryStream.writeUInt32(this.StartSector);
            directoryStream.writeUInt64(this.SizeOfStream);
        }


        // Does nothing in the base class implementation.
        internal virtual void writeReferencedStream()
        {
            return;
        }
    }
}
