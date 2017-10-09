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
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4001)]
    public class StyleTextPropAtom : Record
    {
        public List<ParagraphRun> PRuns;
        public List<CharacterRun> CRuns;

        public StyleTextPropAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.PRuns = new List<ParagraphRun>();
            this.CRuns = new List<CharacterRun>();
        }

        override public void AfterParentSet()
        {
            ClientTextbox textbox = this.ParentRecord as ClientTextbox;
            if (textbox == null)
            {
                TraceLogger.DebugInternal("Record of type StyleTextPropAtom doesn't have parent of type ClientTextbox: {0}", this);
                return;
            }

            // TODO: FindSiblingByType?
            TextAtom textAtom = textbox.FirstChildWithType<TextAtom>();

            /* This can legitimately happen... */
            if (textAtom == null)
            {
                this.Reader.Read(new char[this.BodySize], 0, (int)this.BodySize);
                return;
            }

            // TODO: Length in bytes? UTF-16 characters? Full width unicode characters?

            //TraceLogger.DebugInternal("New text style for text: {0}", Utils.StringInspect(textAtom.Text));

            uint seenLength = 0;
            while (seenLength < textAtom.Text.Length + 1)
            {
                long pos = this.Reader.BaseStream.Position;
                uint length = this.Reader.ReadUInt32();

                ParagraphRun run = new ParagraphRun(this.Reader, false);
                run.Length = length;
                this.PRuns.Add(run);

                /*TraceLogger.DebugInternal("Read paragraph run. Before pos = {0}, after pos = {1} of {2}: {3}",
                    pos, this.Reader.BaseStream.Position, this.Reader.BaseStream.Length,
                    run);*/

                seenLength += length;
            }

            //TraceLogger.DebugInternal();

            seenLength = 0;
            while (seenLength < textAtom.Text.Length + 1)
            {
                uint length = this.Reader.ReadUInt32();

                CharacterRun run = new CharacterRun(this.Reader);
                run.Length = length;
                this.CRuns.Add(run);

                seenLength += length;
            }
        }

        public override string ToString(uint depth)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(base.ToString(depth));

            depth++;
            string indent = IndentationForDepth(depth);

            sb.AppendFormat("\n{0}Paragraph Runs:", indent);
            foreach (ParagraphRun pr in this.PRuns)
                sb.AppendFormat("\n{0}", pr.ToString(depth + 1));

            sb.AppendFormat("\n{0}Character Runs:", indent);
            foreach (CharacterRun cr in this.CRuns)
                sb.AppendFormat("\n{0}", cr.ToString(depth + 1));

            return sb.ToString();
        }
    }

}
