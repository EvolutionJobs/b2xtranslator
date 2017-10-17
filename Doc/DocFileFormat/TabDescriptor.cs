using System;

namespace b2xtranslator.DocFileFormat
{
    public class TabDescriptor
    {
        /// <summary>
        /// Justification code:<br/>
        /// 0 left tab<br/>
        /// 1 centered tab<br/>
        /// 2 right tab<br/>
        /// 3 decimal tab<br/>
        /// 4 bar
        /// </summary>
        public byte jc;

        /// <summary>
        /// Tab leader code:<br/>
        /// 0 no leader<br/>
        /// 1 dotted leader<br/>
        /// 2 hyphenated leader<br/>
        /// 3 single line leader<br/>
        /// 4 heavy line leader<br/>
        /// 5 middle dot
        /// </summary>
        public byte tlc;

        /// <summary>
        /// Parses the bytes to retrieve a TabDescriptor
        /// </summary>
        /// <param name="b">The byte</param>
        public TabDescriptor(byte b)
        {
          this.jc = Convert.ToByte(Convert.ToInt32(b) & 0x07);
          this.tlc = Convert.ToByte(b >> 3);
        }
    }
}
