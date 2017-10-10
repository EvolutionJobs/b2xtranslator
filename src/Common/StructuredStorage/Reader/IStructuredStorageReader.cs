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
namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Reader
{
    public interface IStructuredStorageReader : IDisposable
    {
        /// <summary>
        /// Collection of all _entries contained in a compound file
        /// </summary> 
        ICollection<DirectoryEntry> AllEntries { get; }

        /// <summary> 
        /// Collection of all stream _entries contained in a compound file
        /// </summary> 
        ICollection<DirectoryEntry> AllStreamEntries { get; }
        
        /// <summary>
        /// Collection of all entry names contained in a compound file
        /// </summary>        
        ICollection<string> FullNameOfAllEntries { get; }

        /// <summary>
        /// Collection of all stream entry names contained in a compound file
        /// </summary>        
        ICollection<string> FullNameOfAllStreamEntries { get; }


        DirectoryEntry RootDirectoryEntry { get; }


        /// <summary>
        /// Closes the file handle
        /// </summary>
        void Close();

        /// <summary>
        /// Returns a handle to a stream with the given name/path.
        /// If a path is used, it must be preceeded by '\'.
        /// The characters '\' ( if not separators in the path) and '%' must be masked by '%XXXX'
        /// where 'XXXX' is the unicode in hex of '\' and '%', respectively
        /// </summary>
        /// <param name="path">The path of the virtual stream.</param>
        /// <returns>An object which enables access to the virtual stream.</returns>
        VirtualStream GetStream(string path);
        // TODO: return a System.IO.Stream object only
    }
}
