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


namespace DIaLOGIKa.b2xtranslator.DocFileFormat
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
