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

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public enum SlideSizeType : short
    {
        OnScreen,
        LetterSizedPaper,
        A4Paper,
        Size35mm,
        Overhead,
        Banner,
        Custom
    }

    [OfficeRecord(1001)]
    public class DocumentAtom : Record
    {
        public GPointAtom SlideSize;
        public GPointAtom NotesSize;
        public GRatioAtom ServerZoom;

        public uint NotesMasterPersist;
        public uint HandoutMasterPersist;
        public ushort FirstSlideNum;
        public SlideSizeType SlideSizeType;

        public bool SaveWithFonts;
        public bool OmitTitlePlace;
        public bool RightToLeft;
        public bool ShowComments;

        public DocumentAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.SlideSize = new GPointAtom(this.Reader);
            this.NotesSize = new GPointAtom(this.Reader);
            this.ServerZoom = new GRatioAtom(this.Reader);

            this.NotesMasterPersist = this.Reader.ReadUInt32();
            this.HandoutMasterPersist = this.Reader.ReadUInt32();
            this.FirstSlideNum = this.Reader.ReadUInt16();
            this.SlideSizeType = (SlideSizeType) this.Reader.ReadInt16();

            this.SaveWithFonts = this.Reader.ReadByte() != 0;
            this.OmitTitlePlace = this.Reader.ReadByte() != 0;
            this.RightToLeft = this.Reader.ReadByte() != 0;
            this.ShowComments = this.Reader.ReadByte() != 0;
        }

        override public string ToString(uint depth)
        {
            return String.Format("{0}\n{1}SlideSize = {2}, NotesSize = {3}, ServerZoom = {4}\n{1}" +
                "NotesMasterPersist = {5}, HandoutMasterPersist = {6}, FirstSlideNum = {7}, SlideSizeType = {8}\n{1}" +
                "SaveWithFonts = {9}, OmitTitlePlace = {10}, RightToLeft = {11}, ShowComments = {12}",

                base.ToString(depth), IndentationForDepth(depth + 1),

                this.SlideSize, this.NotesSize, this.ServerZoom,

                this.NotesMasterPersist, this.HandoutMasterPersist, this.FirstSlideNum, this.SlideSizeType,

                this.SaveWithFonts, this.OmitTitlePlace, this.RightToLeft, this.ShowComments);
        }
    }

}
