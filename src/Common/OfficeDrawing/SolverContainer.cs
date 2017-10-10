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
using System.IO;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF005)]
    public class SolverContainer : RegularContainer
    {
        public SolverContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                foreach (var item in Children)
                {
                    switch (item.TypeCode)
                    {
                        default:
                            break;
                    }
                }
        
        }
    }

    [OfficeRecord(0xF012)]
    public class FConnectorRule : Record
    {
        public UInt32 ruid;
        public UInt32 spidA;
        public UInt32 spidB;
        public UInt32 spidC;
        public UInt32 cptiA;
        public UInt32 cptiB;

        public FConnectorRule(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                ruid = this.Reader.ReadUInt32();
                spidA = this.Reader.ReadUInt32();
                spidB = this.Reader.ReadUInt32();
                spidC = this.Reader.ReadUInt32();
                cptiA = this.Reader.ReadUInt32();
                cptiB = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(0xF014)]
    public class FArcRule : Record
    {
        public UInt32 ruid;
        public UInt32 spid;

        public FArcRule(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                ruid = this.Reader.ReadUInt32();
                spid = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(0xF017)]
    public class FCalloutRule : Record
    {
        public UInt32 ruid;
        public UInt32 spid;

        public FCalloutRule(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {
                
                ruid = this.Reader.ReadUInt32();
                spid = this.Reader.ReadUInt32();
        }
    }

}
