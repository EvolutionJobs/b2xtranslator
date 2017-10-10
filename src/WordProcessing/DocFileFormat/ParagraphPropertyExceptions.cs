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
    public class ParagraphPropertyExceptions : PropertyExceptions
    {
        /// <summary>
        /// Index to style descriptor of the style from which the 
        /// paragraph inherits its paragraph and character properties
        /// </summary>
        public ushort istd;

        /// <summary>
        /// Creates a PAPX wich doesn't modify anything.<br/>
        /// The grpprl list is empty
        /// </summary>
        public ParagraphPropertyExceptions() : base()
        {
        }

        /// <summary>
        /// Parses the bytes to retrieve a PAPX
        /// </summary>
        /// <param name="bytes">The bytes starting with the istd</param>
        public ParagraphPropertyExceptions(byte[] bytes, VirtualStream dataStream)
            : base(new List<byte>(bytes).GetRange(2, bytes.Length-2).ToArray())
        {
            if (bytes.Length != 0)
            {
                this.istd = System.BitConverter.ToUInt16(bytes, 0);
            }

            //There is a SPRM that points to an offset in the data stream, 
            //where a list of SPRM is saved.
            foreach (var sprm in this.grpprl)
            {
                if(sprm.OpCode == SinglePropertyModifier.OperationCode.sprmPHugePapx || (int)sprm.OpCode == 0x6646)
                {
                    IStreamReader reader = new VirtualStreamReader(dataStream);
                    var fc = System.BitConverter.ToUInt32(sprm.Arguments, 0);

                    //parse the size of the external grpprl
                    var sizebytes = new byte[2];
                    dataStream.Read(sizebytes, 0, 2, (int)fc);
                    var size = System.BitConverter.ToUInt16(sizebytes, 0);
                    
                    //parse the external grpprl
                    //byte[] grpprlBytes = new byte[size];
                    //dataStream.Read(grpprlBytes);
                    var grpprlBytes = reader.ReadBytes(size);
                    var externalPx = new PropertyExceptions(grpprlBytes);

                    //assign the external grpprl
                    this.grpprl = externalPx.grpprl;

                    //remove the sprmPHugePapx
                    this.grpprl.Remove(sprm);
                }
            }
        }

        #region IVisitable Members

        public override void Convert<T>(T mapping)
        {
            ((IMapping<ParagraphPropertyExceptions>)mapping).Apply(this);
        }

        #endregion
    }
}
