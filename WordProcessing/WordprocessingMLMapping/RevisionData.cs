using System;
using System.Collections.Generic;
using b2xtranslator.DocFileFormat;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class RevisionData
    {
        public enum RevisionType
        {
            NoRevision,
            Inserted,
            Deleted,
            Changed
        }

        public DateAndTime Dttm;
        public short Isbt;
        public RevisionType Type;
        public List<SinglePropertyModifier> Changes;
        public int RsidDel;
        public int RsidProp;
        public int Rsid;

        public RevisionData()
        {
            this.Changes = new List<SinglePropertyModifier>();
        }

        /// <summary>
        /// Collects the revision data of a CHPX
        /// </summary>
        /// <param name="chpx"></param>
        public RevisionData(CharacterPropertyExceptions chpx)
        {
            bool collectRevisionData = true;
            this.Changes = new List<SinglePropertyModifier>();

            foreach (var sprm in chpx.grpprl)
            {
                switch ((int)sprm.OpCode)
                {
                    //revision data
                    case 0xCA89:
                        //revision mark
                        collectRevisionData = false;
                        //author 
                        this.Isbt = System.BitConverter.ToInt16(sprm.Arguments, 1);
                        //date
                        var dttmBytes = new byte[4];
                        Array.Copy(sprm.Arguments, 3, dttmBytes, 0, 4);
                        this.Dttm = new DateAndTime(dttmBytes);
                        break;
                    case 0x0801:
                        //revision mark
                        collectRevisionData = false;
                        break;
                    case 0x4804:
                        //author
                        this.Isbt = System.BitConverter.ToInt16(sprm.Arguments, 0);
                        break;
                    case 0x6805:
                        //date
                        this.Dttm = new DateAndTime(sprm.Arguments);
                        break;
                    case 0x0800:
                        //delete mark
                        this.Type = RevisionType.Deleted;
                        break;
                    case 0x6815:
                        this.RsidProp = System.BitConverter.ToInt32(sprm.Arguments, 0);
                        break;
                    case 0x6816:
                        this.Rsid = System.BitConverter.ToInt32(sprm.Arguments, 0);
                        break;
                    case 0x6817:
                        this.RsidDel = System.BitConverter.ToInt32(sprm.Arguments, 0);
                        break;
                }

                //put the sprm on the revision stack
                if (collectRevisionData)
                {
                    this.Changes.Add(sprm);
                }
            }

            //type
            if (this.Type != RevisionType.Deleted)
            {
                if (collectRevisionData)
                {
                    //no mark was found, so this CHPX doesn't contain revision data
                    this.Type = RevisionType.NoRevision;
                }
                else
                {
                    if (this.Changes.Count > 0)
                    {
                        this.Type = RevisionType.Changed;
                    }
                    else
                    {
                        this.Type = RevisionType.Inserted;
                        this.Changes.Clear();
                    }
                }
            }
        }
    }
}
