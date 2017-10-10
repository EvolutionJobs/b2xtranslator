

using System;
using System.Text;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4086)]
    public class CurrentUserAtom : Record
    {
        /// <summary>
        /// Encoding to use for decoding ANSI strings.
        /// </summary>
        private static readonly Encoding ANSIEncoding = Encoding.GetEncoding("iso-8859-1");

        /// <summary>
        /// An unsigned integer that specifies the length, in bytes, of the fixed-length portion of the record. 
        /// It MUST be 0x00000014.  
        /// </summary>
        public uint Size;

        /// <summary>
        /// An unsigned integer that specifies a token used to identify whether the file is encrypted.
        /// 
        /// It MUST be a value from the following table: 
        ///     0xE391C05F: The file SHOULD NOT be an encrypted document. 
        ///     0xF3D1C4DF: The file MUST be an encrypted document.
        /// </summary>
        public uint HeaderToken;

        /// <summary>
        /// An unsigned integer that specifies an offset, in bytes, from the beginning of the
        /// PowerPoint DocumentRecord Stream to the UserEditAtom record for the most recent user edit. 
        /// </summary>
        public uint OffsetToCurrentEdit;

        /// <summary>
        /// An unsigned integer that specifies the length, in bytes, of the  ansiUserName field.
        /// It MUST be less than or equal to 255. 
        /// </summary>
        public ushort LengthUserName;

        /// <summary>
        /// An unsigned integer that specifies the document file version of the file.
        /// It MUST be 0x03F4. 
        /// </summary>
        public ushort DocFileVersion;

        /// <summary>
        /// An unsigned integer that specifies the major version of the storage format. 
        /// It MUST be 0x03. 
        /// </summary>
        public byte MajorVersion;

        /// <summary>
        /// An unsigned integer that specifies the minor version of the storage format. 
        /// It MUST be 0x00. 
        /// </summary>
        public byte MinorVersion;

        /// <summary>
        /// A PrintableAnsiString that specifies the user name of the last user to
        /// modify the file. The length, in bytes, of the field is specified by the
        /// lenUserName field.  
        /// </summary>
        public string UserNameANSI;

        /// <summary>
        /// An unsigned integer that specifies the release version of the file format.
        /// 
        /// MUST be a value from the following table: 
        ///     0x00000008: The file contains one or more main master slide.
        ///     0x00000009: The file contains more than one main master slide. It SHOULD NOT be used. 
        /// </summary>
        public uint ReleaseVersion;

        /// <summary>
        /// An optional PrintableUnicodeString that specifies the user name of the
        /// last user to modify the file.
        /// 
        /// The length, in bytes, of the field is specified by 2 * lenUserName. 
        /// 
        /// This user name supersedes that specified by the ansiUserName field.
        /// 
        /// It MAY be omitted.
        /// </summary>
        public string UserNameUnicode;

        /// <summary>
        /// UserNameUnicode if present, else UserNameANSI.
        /// </summary>
        public string UserName
        {
            get
            {
                return (this.UserNameUnicode != null) ? this.UserNameUnicode : this.UserNameANSI;
            }
        }

        public CurrentUserAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Size = this.Reader.ReadUInt32();
            this.HeaderToken = this.Reader.ReadUInt32();

            switch (this.HeaderToken)
            {
                case 0xE391C05F: /* regular PPT file */
                    break;

                case 0xF3D1C4DF: /* encrypted PPT file */
                    throw new NotSupportedException("Encryped PPT files aren't supported at this time");

                default:
                    throw new NotSupportedException(string.Format(
                        "File doesn't seem to be a PPT file. Magic Bytes = {0}", this.HeaderToken));
            }

            this.OffsetToCurrentEdit = this.Reader.ReadUInt32();
            this.LengthUserName = this.Reader.ReadUInt16();
            this.DocFileVersion = this.Reader.ReadUInt16();
            this.MajorVersion = this.Reader.ReadByte();
            this.MinorVersion = this.Reader.ReadByte();

            // Throw away reserved data
            this.Reader.ReadUInt16();

            var ansiUserNameBytes = this.Reader.ReadBytes(this.LengthUserName);
            this.UserNameANSI = ANSIEncoding.GetString(ansiUserNameBytes);

            this.ReleaseVersion = this.Reader.ReadUInt32();

            if (this.Reader.BaseStream.Position != this.Reader.BaseStream.Length)
            {
                var unicodeUserNameBytes = this.Reader.ReadBytes(this.LengthUserName * 2);
                this.UserNameUnicode = Encoding.Unicode.GetString(unicodeUserNameBytes);
            }
        }

        override public string ToString(uint depth)
        {
            return string.Format("{0}\n{1}Size = {2}, Magic = {3}, OffsetToCurrentEdit = {4}\n{1}" +
                "LengthUserName = {5}, DocFileVersion = {6}, MajorVersion = {7}, MinorVersion = {8}\n{1}" +
                "UserNameANSI = {9}, ReleaseVersion = {10}, UserNameUnicode = {11}",

                base.ToString(depth), IndentationForDepth(depth + 1),

                this.Size, this.HeaderToken, this.OffsetToCurrentEdit,
                this.LengthUserName, this.DocFileVersion, this.MajorVersion, this.MinorVersion,
                this.UserNameANSI, this.ReleaseVersion, this.UserNameUnicode);
        }
    }

}
