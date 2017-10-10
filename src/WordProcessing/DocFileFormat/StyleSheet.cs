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

using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
{
    public class StyleSheet : IVisitable
    {
        /// <summary>
        /// The StyleSheetInformation of the stylesheet.
        /// </summary>
        public StyleSheetInformation stshi;

        /// <summary>
        /// The list contains all styles.
        /// </summary>
        public List<StyleSheetDescription> Styles;

        /// <summary>
        /// Parses the streams to retrieve a StyleSheet.
        /// </summary>
        /// <param name="fib">The FileInformationBlock</param>
        /// <param name="tableStream">The 0Table or 1Table stream</param>
        public StyleSheet(FileInformationBlock fib, VirtualStream tableStream, VirtualStream dataStream)
        {
            IStreamReader tableReader = new VirtualStreamReader(tableStream);

            //read size of the STSHI
            var stshiLengthBytes = new byte[2];
            tableStream.Read(stshiLengthBytes, 0, stshiLengthBytes.Length, fib.fcStshf);
            var cbStshi = System.BitConverter.ToInt16(stshiLengthBytes, 0);

            //read the bytes of the STSHI
            var stshi = tableReader.ReadBytes(fib.fcStshf + 2, cbStshi);

            //parses STSHI
            this.stshi = new StyleSheetInformation(stshi);

            //create list of STDs
            this.Styles = new List<StyleSheetDescription>();
            for (int i = 0; i < this.stshi.cstd; i++)
            {
                //get the cbStd
                var cbStd = tableReader.ReadUInt16();

                if (cbStd != 0)
                {
                    //read the STD bytes
                    var std = tableReader.ReadBytes(cbStd);

                    //parse the STD bytes
                    this.Styles.Add(new StyleSheetDescription(std, (int)this.stshi.cbSTDBaseInFile, dataStream));
                }
                else
                {
                    this.Styles.Add(null);
                }
            }

        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<StyleSheet>)mapping).Apply(this);
        }

        #endregion
    }
}
