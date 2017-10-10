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
    /// Abstract class for a directory entry in a structured storage.
    /// Athor: math
    /// </summary>
    abstract public class AbstractDirectoryEntry
    {
        uint _sid;
        public uint Sid
        {
            get { return this._sid; }
            internal set { this._sid = value; }
        }

        protected string _path;
        public string Path
        {
            get { return this._path + this.Name; }
        }


        // Name
        protected string _name;
        public string Name
        {
            get { return MaskingHandler.Mask(this._name); }
            protected set {
                this._name = value;
                if (this._name.Length >= 32)
                {
                    throw new InvalidValueInDirectoryEntryException("_ab");                    
                }
            }
        }

        ushort _lengthOfName;
        public ushort LengthOfName
        {
            get
            {
                if (this._name.Length == 0)
                {
                    this._lengthOfName = 0;
                    return 0;
                }

                // length of name in bytes including unicode 0;
                this._lengthOfName = (ushort)((this._name.Length + 1)*2);              
                return this._lengthOfName;
            }
        }



        // Type
        DirectoryEntryType _type;
        public DirectoryEntryType Type
        {
            get { return this._type; }
            protected set
            {
                if ((int)value < 0 || (int)value > 5)
                {
                    throw new InvalidValueInDirectoryEntryException("_mse");
                }
                this._type = value;
            }
        }


        // Color
        DirectoryEntryColor _color;
        public DirectoryEntryColor Color
        {
            get { return this._color; }
            internal set
            {
                if ((int)value < 0 || (int)value > 1)
                {
                    throw new InvalidValueInDirectoryEntryException("_bflags");
                }
                this._color = value;
            }
        }


        // Left sibling sid
        uint _leftSiblingSid;
        public uint LeftSiblingSid
        {
            get { return this._leftSiblingSid; }
            internal set { this._leftSiblingSid = value; }
        }


        // Right sibling sid
        uint _rightSiblingSid;
        public uint RightSiblingSid
        {
            get { return this._rightSiblingSid; }
            internal set { this._rightSiblingSid = value; }
        }


        // Child sibling sid
        uint _childSiblingSid;
        public uint ChildSiblingSid
        {
            get { return this._childSiblingSid; }
            protected set { this._childSiblingSid = value; }
        }


        //CLSID
        Guid _clsId;
        public Guid ClsId
        {
            get { return this._clsId; }
            protected set { this._clsId = value; }
        }


        // User flags
        uint _userFlags;
        public uint UserFlags
        {
            get { return this._userFlags; }
            protected set { this._userFlags = value; }
        }


        // Start sector
        uint _startSector;
        public uint StartSector
        {
            get { return this._startSector; }
            protected set { this._startSector = value; }
        }


        // Size of stream in bytes
        ulong _sizeOfStream;
        public ulong SizeOfStream
        {
            get { return this._sizeOfStream; }
            protected set { this._sizeOfStream = value; }
        }

        internal AbstractDirectoryEntry() : this(0x0)
        {}

        internal AbstractDirectoryEntry(uint sid)
        {
            this._sid = sid;
        }
    }
}
