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
using System.IO;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecord(1000)]
    public class DocumentContainer : RegularContainer
    {
        /// <summary>
        /// Collection of SlidePersistAtoms for master slides.
        /// </summary>
        public List<SlidePersistAtom> MasterPersistList = new List<SlidePersistAtom>();

        /// <summary>
        /// Collection of SlidePersistAtoms for notes slides.
        /// </summary>
        public List<SlidePersistAtom> NotesPersistList = new List<SlidePersistAtom>();

        /// <summary>
        /// Collection of SlidePersistAtoms for regular slides.
        /// </summary>
        public List<SlidePersistAtom> SlidePersistList = new List<SlidePersistAtom>();

        public SlideListWithText RegularSlideListWithText;

        public List DocInfoListContainer;

        public DocumentContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            

            foreach (var collection in this.AllChildrenWithType<SlideListWithText>())
            {
                List<SlidePersistAtom> target = null;

                switch ((SlideListWithText.Instances)collection.Instance)
                {
                    case SlideListWithText.Instances.CollectionOfMasterSlides:
                        target = this.MasterPersistList;
                        break;

                    case SlideListWithText.Instances.CollectionOfNotesSlides:
                        target = this.NotesPersistList;
                        break;

                    case SlideListWithText.Instances.CollectionOfSlides:
                        this.RegularSlideListWithText = collection;
                        target = this.SlidePersistList;
                        break;
                }

                if (target != null)
                {
                    foreach (var atom in collection.AllChildrenWithType<SlidePersistAtom>())
                        target.Add(atom);
                }
            }

            this.MasterPersistList.Sort(delegate(SlidePersistAtom a, SlidePersistAtom b) {
                return a.PersistIdRef.CompareTo(b.PersistIdRef);
            });

            this.NotesPersistList.Sort(delegate(SlidePersistAtom a, SlidePersistAtom b) {
                return a.PersistIdRef.CompareTo(b.PersistIdRef);
            });

            this.SlidePersistList.Sort(delegate(SlidePersistAtom a, SlidePersistAtom b) {
                return a.PersistIdRef.CompareTo(b.PersistIdRef);
            });

            this.DocInfoListContainer = FirstChildWithType<List>();
        }

        public SlidePersistAtom SlidePersistAtomForSlideWithIdx(uint idx)
        {
            foreach (var atom in this.SlidePersistList)
                // idx is zero-based, psr-reference is one-based
                if (atom.PersistIdRef == idx + 1)
                    return atom;

            return null;
        }
    }
}
