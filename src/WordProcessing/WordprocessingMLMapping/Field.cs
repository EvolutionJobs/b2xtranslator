/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System.Collections.Generic;
using System.Text.RegularExpressions;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.DocFileFormat;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
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
