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



/// <summary>
/// Global definitions
/// Author: math
/// </summary>


namespace DIaLOGIKa.b2xtranslator.StructuredStorage.Common
{

    /// <summary>
    /// Constants used to identify sectors in fat, minifat and directory
    /// </summary>
    internal static class SectorId
	{      
        internal const uint MAXREGSECT = 0xFFFFFFFA;
        internal const uint DIFSECT = 0xFFFFFFFC;
        internal const uint FATSECT = 0xFFFFFFFD;
        internal const uint ENDOFCHAIN = 0xFFFFFFFE;
        internal const uint FREESECT = 0xFFFFFFFF;

        internal const uint NOSTREAM = 0xFFFFFFFF;
	}


    /// <summary>
    /// Size constants 
    /// </summary>
    internal static class Measures
    {
        internal const int DirectoryEntrySize = 128;
        internal const int HeaderSize = 512;
    }


    /// <summary>
    /// Type of a directory entry
    /// </summary>
    public enum DirectoryEntryType
    {
        STGTY_INVALID = 0,
        STGTY_STORAGE = 1,
        STGTY_STREAM = 2,
        STGTY_LOCKBYTES = 3,
        STGTY_PROPERTY = 4,
        STGTY_ROOT = 5    
    }


    /// <summary>
    /// Color of a directory entry in the red-black-tree
    /// </summary>
    public enum DirectoryEntryColor
    {
        DE_RED = 0,
        DE_BLACK = 1
    }

}