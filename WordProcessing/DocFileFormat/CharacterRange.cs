namespace b2xtranslator.DocFileFormat
{
    public class CharacterRange
    {
        public int CharacterPosition;
        public int CharacterCount;

        public CharacterRange(int cp, int ccp)
        {
            this.CharacterPosition = cp;
            this.CharacterCount = ccp;
        }
    }
}
