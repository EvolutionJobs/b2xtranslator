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

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class NilPicfAndBinData
    {
        /// <summary>
        /// A signed integer that specifies the size, in bytes, of this structure.
        /// </summary>
        public int lcb;
            
        /// <summary>
        /// An unsigned integer that specifies the number of bytes from the beginning of this structure to the beginning of binData. 
        /// MUST be 0x44. 
        /// </summary>
        public short cbHeader;

        /// <summary>
        /// The interpretation of binData depends on the field type of the field containing the 
        /// picture character and is given by the following table:<br/><br/>
        /// 
        /// REF: HyperlinkFieldData<br/>
        /// PAGEREF: HyperlinkFieldData<br/>
        /// NOTEREF: HyperlinkFieldData<br/><br/>
        /// 
        /// FORMTEXT: FormFieldData<br/>
        /// FORMCHECKBOX: FormFieldData<br/>
        /// FORMDROPDOWN: FormFieldData<br/><br/>
        /// 
        /// PRIVATE: Custom binary data that is specified by the add-in that inserted this field.<br/>
        /// ADDIN: Custom binary data that is specified by the add-in that inserted this field.<br/>
        /// HYPERLINK: HyperlinkFieldData<br/>
        /// </summary>
        public byte[] binData;

        public NilPicfAndBinData(CharacterPropertyExceptions chpx, VirtualStream dataStream)
        {
            //Get start of the NilPicfAndBinData structure
            var fc = PictureDescriptor.GetFcPic(chpx);
            if (fc >= 0)
            {
                parse(dataStream, fc);
            }
        }

        private void parse(VirtualStream stream, int fc)
        {
            stream.Seek(fc, System.IO.SeekOrigin.Begin);
            var reader = new VirtualStreamReader(stream);

            this.lcb = reader.ReadInt32();
            this.cbHeader = reader.ReadInt16();
            reader.ReadBytes(62);
            this.binData = reader.ReadBytes(this.lcb - this.cbHeader);
        }
    }
}
