using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// BOUNDSHEET: Sheet Information (85h)
    /// 
    /// This record stores the sheet name, sheet type, and stream position.
    /// </summary>
    [BiffRecord(RecordType.BoundSheet8)] 
    public class BoundSheet8 : BiffRecord
    {
        public const RecordType ID = RecordType.BoundSheet8;

        public enum HiddenState : byte
        {
            /// <summary>
            /// Visible
            /// </summary>
            Visible = 0x00,
 
            /// <summary>
            /// Hidden
            /// </summary>
            Hidden = 0x01, 

            /// <summary>
            /// Very Hidden; the sheet is hidden and cannot be displayed using the user interface.
            /// </summary>
            VeryHidden = 0x02
        }

        public enum SheetType : ushort
        {
            /// <summary>
            /// Worksheet or dialog sheet
            /// </summary>
            Worksheet = 0x0000,

            /// <summary>
            /// Excel 4.0 macro sheet
            /// </summary>
            Macrosheet = 0x0001,

            /// <summary>
            /// Chart sheet
            /// </summary>
            Chartsheet = 0x0002,

            /// <summary>
            /// Visual Basic module
            /// </summary>
            VisualBasicModule = 0x0006
        } 


        /// <summary>
        /// A FilePointer as specified in [MS-OSHARED] section 2.2.1.5 that specifies the 
        /// stream position of the start of the BOF record for the sheet.
        /// </summary>
        public uint lbPlyPos;

        /// <summary>
        /// A ShortXLUnicodeString that specifies the unique case-insensitive name of the sheet. 
        /// The character count of this string, stName.ch, MUST be greater than or equal to 1 
        /// and less than or equal to 31. 
        /// 
        /// The string MUST NOT contain the any of the following characters: 
        /// 
        ///     - 0x0000 
        ///     - 0x0003 
        ///     - colon (:) 
        ///     - backslash (\) 
        ///     - asterisk (*) 
        ///     - question mark (?) 
        ///     - forward slash (/) 
        ///     - opening square bracket ([) 
        ///     - closing square bracket (]) 
        ///     
        /// The string MUST NOT begin or end with the single quote (') character.
        /// </summary>
        public ShortXLUnicodeString stName;
        // TODO: check for correct interpretation of Unicode strings

        /// <summary>
        /// The hidden status of the workbook 
        /// </summary>
        public HiddenState hsState;

        /// <summary>
        /// The sheet type value
        /// </summary>
        public SheetType dt; 

        /// <summary>
        /// extracts the boundsheetdata from the biffrecord  
        /// </summary>
        /// <param name="reader">IStreamReader </param>
        /// <param name="id">Type of the record </param>
        /// <param name="length">Length of the record</param>
        public BoundSheet8(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            
            this.lbPlyPos = this.Reader.ReadUInt32();

            byte flags = reader.ReadByte();

            // Bitmask is 0003h -> first two bits 
            this.hsState = (HiddenState)Utils.BitmaskToByte(flags, 0x0003); 

            this.dt = (SheetType)reader.ReadByte();

            this.stName = new ShortXLUnicodeString(reader);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }

        /// <summary>
        /// Simple ToString Method 
        /// </summary>
        /// <returns>String from the object</returns>
        public override string ToString()
        {
            string returnvalue = "BOUNDSHEET - RECORD: \n";
            returnvalue += "-- Name: " + this.stName.Value + "\n";
            returnvalue += "-- Offset: " + this.lbPlyPos + "\n";
            returnvalue += "-- HiddenState: " + this.hsState + "\n";
            returnvalue += "-- SheetType: " + this.dt + "\n"; 
            return returnvalue; 
        }
    }
}
