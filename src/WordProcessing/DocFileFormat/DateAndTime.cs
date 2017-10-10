using System;
using System.Collections;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class DateAndTime : IVisitable
    {
        /// <summary>
        /// minutes (0-59)
        /// </summary>
        public short mint;

        /// <summary>
        /// hours (0-23)
        /// </summary>
        public short hr;

        /// <summary>
        /// day of month (1-31)
        /// </summary>
        public short dom;

        /// <summary>
        /// month (1-12)
        /// </summary>
        public short mon;

        /// <summary>
        /// year (1900-2411)-1900
        /// </summary>
        public short yr;

        /// <summary>
        /// weekday<br/>
        /// 0 Sunday
        /// 1 Monday
        /// 2 Tuesday
        /// 3 Wednesday
        /// 4 Thursday
        /// 5 Friday
        /// 6 Saturday
        /// </summary>
        public short wdy;

        /// <summary>
        /// Creates a new DateAndTime with default values
        /// </summary>
        public DateAndTime()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the byte sto retrieve a DateAndTime
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public DateAndTime(byte[] bytes)
        {
            if (bytes.Length == 4)
            {
                var bits = new BitArray(bytes);

                this.mint = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 0, 6));
                this.hr = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 6, 5));
                this.dom = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 11, 5));
                this.mon = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 16, 4));
                this.yr = (short)(1900 + Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 20, 9)));
                this.wdy = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 29, 3));
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct DTTM, the length of the struct doesn't match");
            }
        }

        public DateTime ToDateTime()
        {
            if (this.yr == 1900 && this.mon == 0 && this.dom == 0 && this.hr == 0 && this.mint == 0)
            {
                return new DateTime(1900, 1, 1, 0, 0, 0);
            }
            else
            {
                return new DateTime(this.yr, this.mon, this.dom, this.hr, this.mint, 0);
            } 
        }

        private void setDefaultValues()
        {
            this.dom = 0;
            this.hr = 0;
            this.mint = 0;
            this.mon = 0;
            this.wdy = 0;
            this.yr = 0;
        }

        #region IVisitable Members

        public virtual void Convert<T>(T mapping)
        {
            ((IMapping<DateAndTime>)mapping).Apply(this);
        }

        #endregion
    }
}
