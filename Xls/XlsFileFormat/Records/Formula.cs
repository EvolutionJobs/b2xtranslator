
using System;
using System.Collections.Generic;
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools; 

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Formula)] 
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
        public uint chn;

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
                    this.boolValue = this.val[2];
                    this.boolValueSet = true; 
                }
                if (firstOffset == 2)
                {
                    // this is a error value 
                    this.errorValue = (int)this.val[2];      
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
            this.fAlwaysCalc = Utils.BitmaskToBool((int)this.grbit, 0x01); 

            // check if shared formula
            this.fShrFmla = Utils.BitmaskToBool((int)this.grbit, 0x08);



            oldStreamPosition = this.Reader.BaseStream.Position;
            if (!this.fShrFmla)
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
