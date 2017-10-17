

using System.Collections.Generic;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
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
