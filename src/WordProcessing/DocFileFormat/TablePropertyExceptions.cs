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
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class TablePropertyExceptions : PropertyExceptions
    {
        /// <summary>
        /// Creates a TAPX wich doesn't modify anything.<br/>
        /// The grpprl list is empty
        /// </summary>
        public TablePropertyExceptions()
            : base()
        {
        }

        /// <summary>
        /// Parses the bytes to retrieve a TAPX
        /// </summary>
        /// <param name="bytes">The bytes starting with the istd</param>
        public TablePropertyExceptions(byte[] bytes)
            : base(bytes)
        {
            //not yet implemented
        }

        /// <summary>
        /// Extracts the TAPX SPRMs out of a PAPX
        /// </summary>
        /// <param name="papx"></param>
        public TablePropertyExceptions(ParagraphPropertyExceptions papx, VirtualStream dataStream)
        {
            this.grpprl = new List<SinglePropertyModifier>();
            foreach (var sprm in papx.grpprl)
            {
                if (sprm.Type == SinglePropertyModifier.SprmType.TAP)
                {
                    this.grpprl.Add(sprm);
                }
                else if ((int)sprm.OpCode == 0x646b)
                {
                    IStreamReader reader = new VirtualStreamReader(dataStream);

                    //there is a native TAP in the data stream
                    var fc = System.BitConverter.ToUInt32(sprm.Arguments, 0);
                    
                    //get the size of the following grpprl
                    //byte[] sizebytes = new byte[2];
                    //dataStream.Read(sizebytes, 2, (int)fc);
                    var sizebytes = reader.ReadBytes(fc, 2);
                    var grpprlSize = System.BitConverter.ToUInt16(sizebytes, 0);

                    //read the grpprl
                    //byte[] grpprlBytes = new byte[grpprlSize];
                    //dataStream.Read(grpprlBytes);
                    var grpprlBytes = reader.ReadBytes(grpprlSize);

                    //parse the grpprl
                    var externalPx = new PropertyExceptions(grpprlBytes);

                    foreach (var sprmExternal in externalPx.grpprl)
                    {
                        if (sprmExternal.Type == SinglePropertyModifier.SprmType.TAP)
                        {
                            this.grpprl.Add(sprmExternal);
                        }
                    }
                }
                
            }
        }

        #region IVisitable Members

        public override void Convert<T>(T mapping)
        {
            ((IMapping<TablePropertyExceptions>)mapping).Apply(this);
        }

        #endregion
    }
}
