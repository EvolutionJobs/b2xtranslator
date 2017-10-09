/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using System.Globalization;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public abstract class AbstractPtg
    {
        IStreamReader _reader;
        PtgNumber _id;
        long _offset;
        String data;
        uint length;
        
        protected uint popSize;
        protected PtgType type; 


        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Ptg Id</param>
        /// <param name="length">The recordlength</param>
        public AbstractPtg(IStreamReader reader, PtgNumber ptgid)
        {
            _reader = reader;
            _offset = _reader.BaseStream.Position;
            _id = ptgid;
            this.data = ""; 
        }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Ptg Id</param>
        /// <param name="length">The recordlength</param>
        public AbstractPtg(IStreamReader reader, Ptg0x18Sub ptgid)
        {
            _reader = reader;
            _offset = _reader.BaseStream.Position;
            _id = (PtgNumber)ptgid;
            this.data = "";
        }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Ptg Id</param>
        /// <param name="length">The recordlength</param>
        public AbstractPtg(IStreamReader reader, Ptg0x19Sub ptgid)
        {
            _reader = reader;
            _offset = _reader.BaseStream.Position;
            _id = (PtgNumber)ptgid;
            this.data = "";
        }

        public PtgNumber Id
        {
            get { return _id; }
        }

        public long Offset
        {
            get { return _offset; }
        }

        public IStreamReader Reader
        {
            get { return _reader; }
            set { this._reader = value; }
        }

        protected String Data
        {
            get { return this.data; }
            set { this.data = value; }
        }

        protected uint Length
        {
            get { return this.length; }
            set { this.length = value; }
        }

        public uint getLength()
        {
            return this.length; 
        }

        public String getData()
        {            
            return Convert.ToString(this.data,CultureInfo.GetCultureInfo("en-US"));
        }

        public uint PopSize()
        {
            return this.popSize;
        }

        public PtgType OpType()
        {
            return this.type; 
        }

    }
}
