using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib
{
    public class EmbeddedObjectPart : OpenXmlPart
    {
        public enum ObjectType
        {
            Excel,
            Word,
            Powerpoint,
            Other
        }

        private ObjectType _format;

        public EmbeddedObjectPart(ObjectType format, OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
            _format = format;
        }

        public override string ContentType
        {
            get {
                switch (_format)
                {
                    case ObjectType.Excel:
                        return OpenXmlContentTypes.MSExcel;
                    case ObjectType.Word:
                        return OpenXmlContentTypes.MSWord;
                    case ObjectType.Powerpoint:
                        return OpenXmlContentTypes.MSPowerpoint;
                    case ObjectType.Other:
                        return OpenXmlContentTypes.OleObject;
                    default:
                        return OpenXmlContentTypes.OleObject;
                }
            }
        }

        internal override bool HasDefaultContentType { 
            get {
                return true;
            }         
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.OleObject; }
        }

        public override string TargetName { get { return "oleObject" + this.PartIndex; } }

        //public override string TargetDirectory { get { return "embeddings"; } }

        private string targetdirectory = "embeddings";
        public override string TargetDirectory
        {
            get
            {
                return targetdirectory;
            }

            set
            {
                targetdirectory = value;
            }

        }

        public override string TargetExt 
        { 
            get {
                switch (_format)
                {
                    case ObjectType.Excel:
                        return ".xls";
                    case ObjectType.Word:
                        return ".doc";
                    case ObjectType.Powerpoint:
                        return ".ppt";
                    case ObjectType.Other:
                        return ".bin";
                    default:
                        return ".bin";
                }
            } 
        }
    }
}
