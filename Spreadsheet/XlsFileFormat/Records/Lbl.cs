
using System;
using System.Collections.Generic;
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies a defined name.
    /// </summary>
    /// <remarks>
    /// In the old version of the spec this record has been listed as RecordType.NAME (218h)
    /// whereas in the new version it is listed as RecordType.Lbl (18h)
    /// </remarks>
    [BiffRecord(RecordType.Lbl, RecordType.NAME)]
    public class Lbl : BiffRecord
    {
        public const RecordType ID = RecordType.Lbl;

        public enum FunctionCategory : ushort
        {
            All,
            Financial,
            DateTime,
            MathTrigonometry,
            Statistical,
            Lookup,
            Database,
            Text,
            Logical,
            Info,
            Commands,
            Customize,
            MacroControl,
            DDEExternal,
            UserDefined,
            Engineering,
            Cube
        }
        
        /// <summary>
        /// A bit that specifies whether the defined name is not visible 
        /// in the list of defined names.
        /// </summary>
        public bool fHidden;

        /// <summary>
        /// A bit that specifies whether the defined name represents an XLM macro. 
        /// If this bit is 1, fProc MUST also be 1.
        /// </summary>
        public bool fFunc;

        /// <summary>
        /// A bit that specifies whether the defined name represents a Visual Basic 
        /// for Applications (VBA) macro. If this bit is 1, the fProc MUST also be 1.
        /// </summary>
        public bool fOB;

        /// <summary>
        /// A bit that specifies whether the defined name represents a macro.
        /// </summary>
        public bool fProc;

        /// <summary>
        /// A bit that specifies whether rgce contains a call to a function that can return an array.
        /// </summary>
        public bool fCalcExp;

        /// <summary>
        /// A bit that specifies whether the defined name represents a built-in name.
        /// </summary>
        public bool fBuiltin;

        /// <summary>
        /// An unsigned integer that specifies the function category for the defined name. MUST be less than or equal to 31. 
        /// The values 17 to 31 are user-defined values. User-defined values are specified in FnGroupName. 
        /// The values zero to 16 are defined as specified by the FunctionCategory enum.
        /// </summary>
        public ushort fGrp;

        /// <summary>
        /// A bit that specifies whether the defined name is published. This bit is ignored 
        /// if the fPublishedBookItems field of the BookExt_Conditional12 structure is zero.
        /// </summary>
        public bool fPublished;

        /// <summary>
        /// A bit that specifies whether the defined name is a workbook parameter.
        /// </summary>
        public bool fWorkbookParam;

        /// <summary>
        /// The unsigned integer value of the ASCII character that specifies the shortcut 
        /// key for the macro represented by the defined name. 
        /// 
        /// MUST be zero (No shortcut key) if fFunc is 1 or if fProc is 0. Otherwise MUST 
        /// <84> be greater than or equal to 0x41 and less than or equal to 0x5A, or greater
        /// than or equal to 0x61 and less than or equal to 0x7A.
        /// </summary>
        public byte chKey;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in Name. 
        /// 
        /// MUST be greater than or equal to zero.
        /// </summary>
        public byte cch;

        /// <summary>
        /// An unsigned integer that specifies length of rgce in bytes.
        /// </summary>
        public ushort cce;

        /// <summary>
        /// An unsigned integer that specifies if the defined name is a local name, and if so, 
        /// which sheet it is on. If this is not 0, the defined name is a local name and the value 
        /// MUST be a one-based index to the collection of BoundSheet8 records as they appear in the Global Substream.
        /// </summary>
        public ushort itab;

        /// <summary>
        /// An XLUnicodeStringNoCch that specifies the name for the defined name. If fBuiltin is zero, 
        /// this field MUST satisfy the same restrictions as the name field of XLNameUnicodeString. 
        /// 
        /// If fBuiltin is 1, this field is for a built-in name. Each built-in name has a zero-based 
        /// index value associated to it. A built-in name or its index value MUST be used for this field.
        /// </summary>
        public XLUnicodeStringNoCch Name;
        
        /// <summary>
        /// A NameParsedFormula that specifies the formula for the defined name.
        /// </summary>
        public Stack<AbstractPtg> rgce;

        public Lbl(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            //Debug.Assert(this.Id == ID);

            ushort flags = reader.ReadUInt16();
            this.fHidden = Utils.BitmaskToBool(flags, 0x0001);
            this.fFunc = Utils.BitmaskToBool(flags, 0x0002);
            this.fOB = Utils.BitmaskToBool(flags, 0x0004);
            this.fProc = Utils.BitmaskToBool(flags, 0x0008);
            this.fCalcExp = Utils.BitmaskToBool(flags, 0x0010);
            this.fBuiltin = Utils.BitmaskToBool(flags, 0x0020);
            this.fGrp = Utils.BitmaskToUInt16(flags, 0x07C0);
            this.fPublished = Utils.BitmaskToBool(flags, 0x2000);
            this.fWorkbookParam = Utils.BitmaskToBool(flags, 0x4000);

            this.chKey = reader.ReadByte();
            this.cch = reader.ReadByte();
            this.cce = reader.ReadUInt16();
            //read reserved bytes 
            reader.ReadBytes(2);
            this.itab = reader.ReadUInt16();
            // read 4 reserved bytes 
            reader.ReadBytes(4);

            if (this.cch > 0)
            {
                this.Name = new XLUnicodeStringNoCch(reader, this.cch);
            }
            else
            {
                this.Name = new XLUnicodeStringNoCch();
            }
            long oldStreamPosition = this.Reader.BaseStream.Position;
            try
            {
                this.rgce = ExcelHelperClass.getFormulaStack(this.Reader, this.cce);
            }
            catch (Exception ex)
            {
                this.Reader.BaseStream.Seek(oldStreamPosition, System.IO.SeekOrigin.Begin);
                this.Reader.BaseStream.Seek(this.cce, System.IO.SeekOrigin.Current);
                TraceLogger.Error("Formula parse error in intern name");
                TraceLogger.Debug(ex.StackTrace);
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
