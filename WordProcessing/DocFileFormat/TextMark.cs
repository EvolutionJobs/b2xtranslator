namespace b2xtranslator.DocFileFormat
{
    public class TextMark
    {
        public const char ParagraphEnd = (char)13;
        public const char HardLineBreak = (char)11;
        public const char BreakingHyphen = (char)4;
        public const char NonRequiredHyphen = (char)31;
        public const char NonBreakingHyphen = (char)30;
        public const char NonBreakingSpace = (char)160;
        public const char Space = (char)32;
        public const char PageBreakOrSectionMark = (char)12;
        public const char ColumnBreak = (char)14;
        public const char Tab = (char)9;
        public const char FieldBeginMark = (char)19;
        public const char FieldEndMark = (char)21;
        public const char FieldSeperator = (char)20;
        public const char CellOrRowMark = (char)7;

        //Special Characters (chp.fSpec == 1)

        public const char CurrentPageNumber = (char)0;
        public const char Picture = (char)1;
        public const char AutoNumberedFootnoteReference = (char)2;
        public const char FootnoteSeparator = (char)3;
        public const char FootnoteContinuation = (char)4;
        public const char AnnotationReference = (char)5;
        public const char LineNumber = (char)6;
        public const char HandAnnotationPicture = (char)7;
        public const char DrawnObject = (char)8;
        public const char Symbol = (char)40;
    }
}
