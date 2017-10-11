namespace b2xtranslator.DocFileFormat
{
    public class PictureBulletInformation
    {
        public bool fPicBullet;
        public bool fNoAutoSize;
        public bool fDefaultPic;
        public bool fTemporary;
        public bool fFormatting;
        public int iBullet;

        public PictureBulletInformation(byte[] bytes)
        {
            if (bytes.Length == 6)
            {
                short flag = System.BitConverter.ToInt16(bytes, 0);
                this.fPicBullet = Tools.Utils.BitmaskToBool(flag, 0x0001);
                this.fNoAutoSize = Tools.Utils.BitmaskToBool(flag, 0x0002);
                this.fDefaultPic = Tools.Utils.BitmaskToBool(flag, 0x0004);
                this.fTemporary = Tools.Utils.BitmaskToBool(flag, 0x0008);
                this.fFormatting = Tools.Utils.BitmaskToBool(flag, 0x1000);
                this.iBullet = System.BitConverter.ToInt32(bytes, 2);
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct PBI, the length of the struct doesn't match");
            }
        }
    }
}
