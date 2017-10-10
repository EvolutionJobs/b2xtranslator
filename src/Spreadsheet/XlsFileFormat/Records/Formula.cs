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
using System.Collections.Generic;
using System.Diagnostics;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools; 

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecordAttribute(RecordType.Formula)] 
    public class Formula : BiffRecord
    {
        public const RecordType ID = RecordType.Formula;
        /// <summary>
        /// Rownumber 
        /// </summary>
        public ushort rw;

        /// <summary>
        /// Colnumber 
        /// </summary>
        public ushort col;

        /// <summary>
        /// index to the XF record 
        /// </summary>
        public ushort ixfe;

        /// <summary>
        /// 8 byte calculated number / string / error of the formular 
        /// </summary>
        public byte[] val;

        /// <summary>
        /// option flags 
        /// </summary>
        public ushort grbit;

        /// <summary>
        /// used for performance reasons only 
        /// can be ignored 
        /// </summary>
        public UInt32 chn;

        /// <summary>
        /// length of the formular data !!
        /// </summary>
        public ushort cce;

        /// <summary>
        /// this attribute indicates if the formula is a shared formula 
        /// </summary>
        public Boolean fShrFmla;

        /// <summary>
        /// LinkedList with the Ptg records !!
        /// </summary>
        public Stack<AbstractPtg> ptgStack;

        /// <summary>
        ///  this is the calculated value 
        /// </summary>
        public double calculatedValue;
        public bool boolValueSet;
        public byte boolValue; 
        public int errorValue;
        public bool fAlwaysCalc; 

        public Formula(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.val = new byte[8];
            this.rw = reader.ReadUInt16();
            this.col = reader.ReadUInt16();
            this.ixfe = reader.ReadUInt16();
            this.boolValueSet = false;

            long oldStreamPosition = this.Reader.BaseStream.Position;
            this.val = reader.ReadBytes(8); // read 8 bytes for the value of the formular            
            if (this.val[6] == 0xFF && this.val[7] == 0xFF)
            {
                // this value is a string, an error or a boolean value
                byte firstOffset = this.val[0];
                if (firstOffset == 1)
                {
                    // this is a boolean value 
                    this.boolValue = val[2];
                    this.boolValueSet = true; 
                }
                if (firstOffset == 2)
                {
                    // this is a error value 
                    this.errorValue = (int) val[2];      
                }
            }
            else
            {
                this.Reader.BaseStream.Seek(oldStreamPosition, System.IO.SeekOrigin.Begin);
                this.calculatedValue = reader.ReadDouble(); 
            }


            this.grbit = reader.ReadUInt16();
            this.chn = reader.ReadUInt32(); // this is used for performance reasons only 
            this.cce = reader.ReadUInt16();
            this.ptgStack = new Stack<AbstractPtg>();
            // reader.ReadBytes(this.cce);

            // check always calc mode 
            this.fAlwaysCalc = Utils.BitmaskToBool((int)grbit, 0x01); 

            // check if shared formula
            this.fShrFmla = Utils.BitmaskToBool((int)grbit, 0x08);



            oldStreamPosition = this.Reader.BaseStream.Position;
            if (!fShrFmla)
            {
                try
                {
                    this.ptgStack = ExcelHelperClass.getFormulaStack(this.Reader, this.cce);
                }
                catch (Exception ex)
                {
                    this.Reader.BaseStream.Seek(oldStreamPosition, System.IO.SeekOrigin.Begin);
                    this.Reader.BaseStream.Seek(this.cce, System.IO.SeekOrigin.Current);
                    TraceLogger.Error("Formula parse error in Row {0} Column {1}", this.rw, this.col);
                    TraceLogger.Debug(ex.StackTrace);
                    TraceLogger.Debug("Inner exception: {0}", ex.InnerException.StackTrace);
                }
            }
            else
            {
                reader.ReadBytes(this.cce); 
            }
           
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }


        public override string ToString()
        {
            return "Fomula at position: Row - " + this.rw.ToString() + " | Col - " + this.col.ToString();   
        }


    }
}
