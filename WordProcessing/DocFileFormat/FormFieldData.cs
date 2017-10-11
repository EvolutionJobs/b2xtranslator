using b2xtranslator.Tools;
using b2xtranslator.CommonTranslatorLib;

namespace b2xtranslator.DocFileFormat
{
    public class FormFieldData : IVisitable
    {
        /// <summary>
        /// An unsigned integer that MUST be 0xFFFFFFFF.
        /// </summary>
        public uint version;

        /// <summary>
        /// Specifies the type of the form field.
        /// </summary>
        public FormFieldType iType;

        /// <summary>
        /// An unsigned integer.
        /// If iType is iTypeText (0), then iRes MUST be 0.<br/>
        /// If iType is iTypeChck (1), then iRes specifies the state of the checkbox and 
        /// MUST be 0 (unchecked), 1 (checked), or 25 (undefined).<br/> 
        /// Undefined checkboxes are treated as unchecked.<br/> 
        /// If iType is iTypeDrop (2), then iRes specifies the current selected list box item.<br/>
        /// A value of 25 specifies the selection is undefined. 
        /// Otherwise, iRes is a zero-based index into FFData.hsttbDropList.
        /// </summary>
        public ushort iRes;

        /// <summary>
        /// A bool that specifies whether the form field has custom help text in xstzHelpText. <br/>
        /// If fOwnHelp is false, then xstzHelpText contains an empty or auto-generated string.
        /// </summary>
        public bool fOwnHelp;

        /// <summary>
        /// A bool that specifies whether the form field has custom status bar text in xstzStatText. <br/>
        /// If fOwnStat is false, then xstzStatText contains an empty or auto-generated string.
        /// </summary>
        public bool fOwnStat;

        /// <summary>
        /// A bool that specifies whether the form field is protected and its value cannot be changed.
        /// </summary>
        public bool fProt;

        /// <summary>
        /// A bit that specifies whether a checkbox‘s size is automatically determined 
        /// by the text size where the checkbox is located. <br/><br/>
        /// 
        /// MUST be 0 if iType is not iTypeChck (1).
        /// </summary>
        public byte iSize;

        /// <summary>
        /// Specifies the type of the textbox
        /// </summary>
        public TextboxType iTypeTxt;

        /// <summary>
        /// A bool that specifies whether the field‘s value is automatically calculated after the field is modified.
        /// </summary>
        public bool fRecalc;

        /// <summary>
        /// A bool that specifies that the form field has a list box. <br/<br/>
        /// 
        /// MUST be true if iType is iTypeDrop. 
        /// Otherwise, MUST be false.
        /// </summary>
        public bool fHasListBox;

        /// <summary>
        /// An unsigned integer that specifies the maximum length, in characters, 
        /// of the value of the textbox.<br/><br/>
        /// 
        /// MUST NOT exceed 32767.<br/> 
        /// A value of 0 means there is no maximum length of the value of the textbox.<br/> 
        /// MUST be 0 if iType is not iTypeText (0).<br/>
        /// </summary>
        public ushort cch;

        /// <summary>
        /// An unsigned integer.<br/><br/>
        /// 
        /// If iType is iTypeChck (1), then hps specifies the size, in half-points, 
        /// of the checkbox and MUST be between 2 and 3168, inclusive.<br/> 
        /// If bitiType is not iTypeChck (1), then hps is undefined and MUST be ignored.
        /// </summary>
        public ushort hps;

        /// <summary>
        /// An string that specifies the name of this form field.<br/><br/>
        /// 
        /// The length MUST NOT exceed 20.
        /// </summary>
        public string xstzName;

        /// <summary>
        /// An optional Xstz that specifies the default text of this textbox.<br/><br/>
        /// 
        /// This structure MUST exist if and only if iType is iTypeTxt (0).<br/> 
        /// The length MUST NOT exceed 255.<br/>
        /// If iTypeTxt is either iTypeTxtCurDate (3) or iTypeTxtCurTime (4), 
        /// then xstzTextDef MUST be an empty string.<br/> 
        /// If iTypeTxt is iTypeTxtCalc (5), then xstzTextDef specifies an expression to calculate.
        /// </summary>
        public string xstzTextDef;

        /// <summary>
        /// An optional unsigned integer that specifies the default state of the checkbox or dropdown list box.<br/><br/>
        /// 
        /// MUST exist if and only if iType is iTypeChck (1) or iTypeDrop (2).<br/> 
        /// If iType is iTypeChck (1), then wDef MUST be 0 or 1 and specify 
        /// the default state of the checkbox as unchecked or checked, respectively.<br/> 
        /// If iType is iTypeDrop (2), then wDef MUST be less than the number of 
        /// items in the dropdown list box and specify the default item selected (zero-based index).
        /// </summary>
        public ushort wDef;

        /// <summary>
        /// An string that specifies the string format of the textbox. <br/><br/>
        /// 
        /// MUST be an empty string if iType is not iTypeTxt (0).<br/> 
        /// The length MUST NOT exceed 64.<br/> 
        /// Valid formatting strings are specified in [ECMA-376] part 4, section 2.16.22 format (Text Box Form Field Formatting).
        /// </summary>
        public string xstzTextFormat;

        /// <summary>
        /// An string that specifies the help text for the form field.<br/><br/>
        /// 
        /// The length MUST NOT exceed 255.
        /// </summary>
        public string xstzHelpText;

        /// <summary>
        /// An string that specifies the status bar text for the form field.<br/><br/>
        /// 
        /// The length MUST NOT exceed 138.
        /// </summary>
        public string xstzStatText;

        /// <summary>
        /// An string that specifies a macro to run upon entry of the form field.<br/><br/>
        /// 
        /// The length MUST NOT exceed 32.
        /// </summary>
        public string xstzEntryMcr;

        /// <summary>
        /// An string that specifies a macro to run after the value of the form field has changed. <br/><br/>
        /// 
        /// The length MUST NOT exceed 32.
        /// </summary>
        public string xstzExitMcr;

        /// <summary>
        /// An optional STTB that specifies the _entries in the dropdown list box. <br/><br/>
        /// 
        /// MUST exist if and only if iType is iTypeDrop (2). 
        /// Entries are Unicode strings and do not have extra data. 
        /// MUST NOT exceed 25 elements.
        /// </summary>
        public string[] hsttbDropList;

        /// <summary>
        /// Creates a new FFData by reading the data from the given stream.<br/>
        /// The position must already be set.
        /// </summary>
        /// <param name="dataStream"></param>
        public FormFieldData(byte[] bytes)
        {
            int pos = 0;
            this.version = System.BitConverter.ToUInt32(bytes, pos);

            if (this.version == 0xFFFFFFFF)
            {
                pos += 4;

                int bits = (int)System.BitConverter.ToUInt16(bytes, pos);
                this.iType = (FormFieldType)Utils.BitmaskToInt(bits, 0x3);
                this.iRes = (ushort)Utils.BitmaskToInt(bits, 0x7C);
                this.fOwnHelp = Utils.BitmaskToBool(bits, 0x80);
                this.fOwnStat = Utils.BitmaskToBool(bits, 0x100);
                this.fProt = Utils.BitmaskToBool(bits, 0x200);
                this.iSize = (byte)Utils.BitmaskToInt(bits, 0x400);
                this.iTypeTxt = (TextboxType)Utils.BitmaskToInt(bits, 0x3800);
                this.fRecalc = Utils.BitmaskToBool(bits, 0x4000);
                this.fHasListBox = Utils.BitmaskToBool(bits, 0x8000);
                pos += 2;

                this.cch = System.BitConverter.ToUInt16(bytes, pos);
                pos += 2;

                this.hps = System.BitConverter.ToUInt16(bytes, pos);
                pos += 2;

                //read the name
                this.xstzName = Utils.ReadXstz(bytes, pos);
                pos += (this.xstzName.Length * 2) + 2 + 2;

                //read text def
                if (this.iType == FormFieldType.iTypeText)
                {
                    this.xstzTextDef = Utils.ReadXstz(bytes, pos);
                    pos += (this.xstzTextDef.Length * 2) + 2 + 2;
                }

                //definition
                if (this.iType == FormFieldType.iTypeChck || this.iType == FormFieldType.iTypeDrop)
                {
                    this.wDef = System.BitConverter.ToUInt16(bytes, pos);
                    pos += 2;
                }

                //read the text format
                this.xstzTextFormat = Utils.ReadXstz(bytes, pos);
                pos += (this.xstzTextFormat.Length * 2) + 2 + 2;

                //read the help test
                this.xstzHelpText = Utils.ReadXstz(bytes, pos);
                pos += (this.xstzHelpText.Length * 2) + 2 + 2;

                //read the status
                this.xstzStatText = Utils.ReadXstz(bytes, pos);
                pos += (this.xstzStatText.Length * 2) + 2 + 2;

                //read the entry macro
                this.xstzEntryMcr = Utils.ReadXstz(bytes, pos);
                pos += (this.xstzEntryMcr.Length * 2) + 2 + 2;

                //read the exit macro
                this.xstzExitMcr = Utils.ReadXstz(bytes, pos);
                pos += (this.xstzExitMcr.Length * 2) + 2 + 2;
            }
        }

        public enum FormFieldType
        {
            /// <summary>
            /// Specifies that the form field is a textbox.
            /// </summary>
            iTypeText,

            /// <summary>
            /// Specifies that the form field is a checkbox.
            /// </summary>
            iTypeChck,

            /// <summary>
            /// Specifies that the form field is a dropdown list box.
            /// </summary>
            iTypeDrop
        }

        public enum TextboxType
        { 
            /// <summary>
            /// Specifies that the textbox value is regular text.
            /// </summary>
            regular,

            /// <summary>
            /// Specifies that the textbox value is a number.
            /// </summary>
            number,

            /// <summary>
            /// Specifies that the textbox value is a date or time.
            /// </summary>
            date,

            /// <summary>
            /// Specifies that the textbox value is the current date.
            /// </summary>
            currentDate,

            /// <summary>
            /// Specifies that the textbox value is the current time.
            /// </summary>
            currentTime,

            /// <summary>
            /// Specifies that the textbox value is calculated from an expression. 
            /// The expression is given by xstzTextDef.
            /// </summary>
            calculated
        }

        #region IVisitable Members

        public virtual void Convert<T>(T mapping)
        {
            ((IMapping<FormFieldData>)mapping).Apply(this);
        }

        #endregion
    }
}
