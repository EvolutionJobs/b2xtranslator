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


using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Writer
{

    /// <summary>
    /// Class which pools the different elements of a structured storage in a context.
    /// Author math.
    /// </summary>
    internal class StructuredStorageContext
    {
        private uint _sidCounter = 0x0;

        // The header of this context.
        Header _header;
        internal Header Header
        {
            get { return this._header; }            
        }

        // The fat of this context.
        Fat _fat;
        internal Fat Fat
        {
            get { return this._fat; }            
        }

        // The mini fat of this context.
        MiniFat _miniFat;
        internal MiniFat MiniFat
        {
            get { return this._miniFat; }            
        }

        // The handler of the output stream of this context.
        OutputHandler _tempOutputStream;
        internal OutputHandler TempOutputStream
        {
            get { return this._tempOutputStream; }            
        }

        // The handler of the directory stream of this context.
        OutputHandler _directoryStream;
        internal OutputHandler DirectoryStream
        {
            get { return this._directoryStream; }
        }

        // The internal bit converter of this context.
        InternalBitConverter _internalBitConverter;
        internal InternalBitConverter InternalBitConverter
        {
            get { return this._internalBitConverter; }
        }

        // The root directroy entry of this context.
        private RootDirectoryEntry _rootDirectoryEntry;
        public RootDirectoryEntry RootDirectoryEntry
        {
            get { return this._rootDirectoryEntry; }
        }


        /// <summary>
        /// Constructor.
        /// </summary>
        internal StructuredStorageContext()
        {
            this._tempOutputStream = new OutputHandler(new MemoryStream());
            this._directoryStream = new OutputHandler(new MemoryStream());
            this._header = new Header(this);
            this._internalBitConverter = new InternalBitConverter(true);
            this._fat = new Fat(this);
            this._miniFat = new MiniFat(this);
            this._rootDirectoryEntry = new RootDirectoryEntry(this);
        }


        /// <summary>
        ///  Returns a new sid for directory entries in this context.
        /// </summary>
        /// <returns>The new sid.</returns>
        internal uint getNewSid()
        {
            return ++this._sidCounter;
        }
    }
}
