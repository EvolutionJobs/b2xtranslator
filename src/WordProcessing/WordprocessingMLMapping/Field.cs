using System.Collections.Generic;
using System.Text.RegularExpressions;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class Field : IVisitable
    {
        public string FieldCode;

        public string FieldExpansion;

        private Regex classicFieldFormat = new Regex(@"^(" + TextMark.FieldBeginMark + ")(.*)(" + TextMark.FieldSeperator + ")(.*)(" + TextMark.FieldEndMark + ")");
        
        private Regex shortFieldFormat = new Regex(@"^(" + TextMark.FieldBeginMark + ")(.*)(" + TextMark.FieldEndMark + ")");

        public Field(char[] fieldChars)
        {
            parse(new string(fieldChars));
        }

        public Field(List<char> fieldChars)
        {
            parse(new string(fieldChars.ToArray()));
        }

        public Field(string fieldString)
        {
            parse(fieldString);
        }

        private void parse(string field)
        {
            if (this.classicFieldFormat.IsMatch(field))
            {
                var classic = this.classicFieldFormat.Match(field);
                this.FieldCode = classic.Groups[2].Value;
                this.FieldExpansion = classic.Groups[4].Value;
            }
            else if (this.shortFieldFormat.IsMatch(field))
            {
                var shortField = this.shortFieldFormat.Match(field);
                this.FieldCode = shortField.Groups[2].Value;
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<Field>)mapping).Apply(this);
        }

        #endregion
    }
}
