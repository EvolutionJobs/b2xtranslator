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
using System.IO;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Common;

[assembly: CLSCompliant(false)]

namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Reader
{

    /// <summary>
    /// Provides methods for accessing a compound file.
    /// Author: math
    /// </summary>
    public sealed class StructuredStorageReader : 
        IStructuredStorageReader
    {

        InputHandler _fileHandler;
        Header _header;
        Fat _fat;
        MiniFat _miniFat;
        DirectoryTree _directory;

        /// <summary>Get a collection of all entry names contained in a compound file</summary>        
        public ICollection<string> FullNameOfAllEntries => this._directory.GetPathsOfAllEntries();

        /// <summary>Get a collection of all stream entry names contained in a compound file</summary>        
        public ICollection<string> FullNameOfAllStreamEntries => this._directory.GetPathsOfAllStreamEntries();

        /// <summary>Get a collection of all _entries contained in a compound file</summary> 
        public ICollection<DirectoryEntry> AllEntries => this._directory.GetAllEntries();

        /// <summary>Get a collection of all stream _entries contained in a compound file</summary> 
        public ICollection<DirectoryEntry> AllStreamEntries => this._directory.GetAllStreamEntries();

        /// <summary>Get a handle to the RootDirectoryEntry.</summary>
        public DirectoryEntry RootDirectoryEntry => this._directory.GetDirectoryEntry(0x0);

        /// <summary>Initalizes a handle to a compound file based on a stream</summary>
        /// <param name="stream">The stream to the storage</param>
        public StructuredStorageReader(Stream stream)
        {
            try
            {
                this._fileHandler = new InputHandler(stream);
                this._header = new Header(this._fileHandler);
                this._fat = new Fat(this._header, this._fileHandler);
                this._directory = new DirectoryTree(this._fat, this._header, this._fileHandler);
                this._miniFat = new MiniFat(this._fat, this._header, this._fileHandler, this._directory.GetMiniStreamStart(), this._directory.GetSizeOfMiniStream());
            }
            catch
            {
                this.Close();
                throw;
            }
        }

        /// <summary>Initalizes a handle to a compound file with the given name</summary>
        /// <param name="fileName">The name of the file including its path</param>
        public StructuredStorageReader(string fileName)
        {
            try
            {
                this._fileHandler = new InputHandler(fileName);
                this._header = new Header(this._fileHandler);
                this._fat = new Fat(this._header, this._fileHandler);
                this._directory = new DirectoryTree(this._fat, this._header, this._fileHandler);
                this._miniFat = new MiniFat(this._fat, this._header, this._fileHandler, this._directory.GetMiniStreamStart(), this._directory.GetSizeOfMiniStream());
            }
            catch
            {
                this.Close();
                throw;
            }
        }


        /// <summary>
        /// Returns a handle to a stream with the given name/path.
        /// If a path is used, it must be preceeded by '\'.
        /// The characters '\' ( if not separators in the path) and '%' must be masked by '%XXXX'
        /// where 'XXXX' is the unicode in hex of '\' and '%', respectively
        /// </summary>
        /// <param name="path">The path of the virtual stream.</param>
        /// <returns>An object which enables access to the virtual stream.</returns>
        public VirtualStream GetStream(string path)
        {
            var entry = this._directory.GetDirectoryEntry(path);
            if (entry == null)
                throw new StreamNotFoundException(path);

            if (entry.Type != DirectoryEntryType.STGTY_STREAM)
                throw new WrongDirectoryEntryTypeException();

            // only streams up to long.MaxValue are supported
            if (entry.SizeOfStream > long.MaxValue)
                throw new UnsupportedSizeException(entry.SizeOfStream.ToString());

            // Determine whether this stream is a "normal stream" or a stream in the mini stream
            if (entry.SizeOfStream < this._header.MiniSectorCutoff)
                return new VirtualStream(this._miniFat, entry.StartSector, (long)entry.SizeOfStream, path);

            return new VirtualStream(this._fat, entry.StartSector, (long)entry.SizeOfStream, path);
        }

        /// <summary>
        /// Returns a handle to a directory entry with the given name/path.
        /// If a path is used, it must be preceeded by '\'.
        /// The characters '\' ( if not separators in the path) and '%' must be masked by '%XXXX'
        /// where 'XXXX' is the unicode in hex of '\' and '%', respectively
        /// </summary>
        /// <param name="path">The path of the directory entry.</param>
        /// <returns>An object which enables access to the directory entry.</returns>
        public DirectoryEntry GetEntry(string path)
        {
            var entry = this._directory.GetDirectoryEntry(path);
            if (entry == null)
                throw new DirectoryEntryNotFoundException(path);

            return entry;
        }


        /// <summary>Closes the file handle</summary>
        public void Close() =>this._fileHandler?.CloseStream();
       

        public void Dispose() => this.Close();
    }
}