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
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// BOOKBOOL: Workbook Option Flag (DAh)
    /// 
    /// This record saves a workbook option flag.
    /// </summary>
    [BiffRecordAttribute(RecordType.BookBool)] 
    public class BookBool : BiffRecord
    {
        public const RecordType ID = RecordType.BookBool;

        /// <summary>
        /// An option flag. See other members.
        /// </summary>
        public ushort grbit;

        //The grbit field contains the following flags:
        //Bits	Mask	Flag Name	Contents
	    public bool fNoSaveSupp;  	    //  0	    0001h    =1 if the Save External Link Values option is turned off (Options dialog box, Calculation tab) 
	    public bool reserved0;	        //  1	    0002h
	    public bool fHasEnvelope; 	    //  2	    0004h    xl9:   =1 if book has envelope (File | Send To | Mail Recipient ) 
	    public bool fEnvelopeVisible;   // 	3	    0008h    xl9:   =1 if envelope is visible
	    public bool fEnvelopeInitDone; 	//  4	    0010h    xl10:  =1 if envelope has been initialized
	    public uint grbitUpdateLinks; //  6-5	    0060h    xl10: Update external links:
                                        //                          0= prompt user to update
                                        //                          1= do not prompt, do not update
                                        //                          2= do not prompt, do update 
	    public bool reserved1;	        //  7	    0080h
	    public bool fHideBorderUnsels; 	//  8	    0100h    xl11:  1= hide borders of unselected Tables 
	    public uint reserved2;	    //  15-9	FE00h


        public BookBool(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            grbit = reader.ReadUInt16();

            fNoSaveSupp = Utils.BitmaskToBool(grbit, 0x0001);  	  
            reserved0 = Utils.BitmaskToBool(grbit, 0x0002);	      
            fHasEnvelope = Utils.BitmaskToBool(grbit, 0x0004); 	  
            fEnvelopeVisible = Utils.BitmaskToBool(grbit, 0x0008); 
            fEnvelopeInitDone = Utils.BitmaskToBool(grbit, 0x0010);
            grbitUpdateLinks = (uint)Utils.BitmaskToInt(grbit, 0x0060); 
            reserved1 = Utils.BitmaskToBool(grbit, 0x0080);	      
            fHideBorderUnsels = Utils.BitmaskToBool(grbit, 0x0100);
            reserved2 = (uint)Utils.BitmaskToInt(grbit, 0xFE00);	  

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
