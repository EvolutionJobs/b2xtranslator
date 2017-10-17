

using System;
using System.Collections.Generic;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4080)]
    public class SlideListWithText : RegularContainer
    {
        public enum Instances
        {
            CollectionOfSlides = 0,
            CollectionOfMasterSlides = 1,
            CollectionOfNotesSlides = 2
        };

        /// <summary>
        /// List of all SlidePersistAtoms of this SlideListWithText.
        /// </summary>
        public List<SlidePersistAtom> SlidePersistAtoms = new List<SlidePersistAtom>();

        /// <summary>
        /// This dictionary manages a list of TextHeaderAtoms for each SlidePersistAtom.
        /// 
        /// Text of placeholders can appear in the SlideListWithText record.
        /// This dictionary is used for associating such text records with the slide they appear on.
        /// </summary>
        public Dictionary<SlidePersistAtom, List<TextHeaderAtom>> SlideToPlaceholderTextHeaders =
            new Dictionary<SlidePersistAtom,List<TextHeaderAtom>>();

        public Dictionary<SlidePersistAtom, List<TextSpecialInfoAtom>> SlideToPlaceholderSpecialInfo =
           new Dictionary<SlidePersistAtom, List<TextSpecialInfoAtom>>();

        public SlideListWithText(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            SlidePersistAtom curSpAtom = null;
            TextHeaderAtom curThAtom = null;

            foreach (var r in this.Children)
            {
                var spAtom = r as SlidePersistAtom;
                var thAtom = r as TextHeaderAtom;
                var tdRecord = r as ITextDataRecord;
                var tsiAtom = r as TextSpecialInfoAtom;

                if (spAtom != null)
                {
                    curSpAtom = spAtom;
                    this.SlidePersistAtoms.Add(spAtom);
                }

                else if (thAtom != null)
                {
                    curThAtom = thAtom;

                    if (!this.SlideToPlaceholderTextHeaders.ContainsKey(curSpAtom))
                        this.SlideToPlaceholderTextHeaders[curSpAtom] = new List<TextHeaderAtom>();

                    this.SlideToPlaceholderTextHeaders[curSpAtom].Add(thAtom);
                }

                else if (tdRecord != null)
                {
                    curThAtom.HandleTextDataRecord(tdRecord);
                }
                else if (tsiAtom != null)
                {
                    if (!this.SlideToPlaceholderSpecialInfo.ContainsKey(curSpAtom))
                        this.SlideToPlaceholderSpecialInfo[curSpAtom] = new List<TextSpecialInfoAtom>();

                    this.SlideToPlaceholderSpecialInfo[curSpAtom].Add(tsiAtom);
                }
                else
                {

                }
            }
        }

        public TextHeaderAtom FindTextHeaderForOutlineTextRef(OutlineTextRefAtom otrAtom)
        {
            var slide = otrAtom.FirstAncestorWithType<Slide>();

            if (slide == null)
                throw new NotSupportedException("Can't find TextHeaderAtom for OutlineTextRefAtom which has no Slide ancestor");

            var thAtoms = this.SlideToPlaceholderTextHeaders[slide.PersistAtom];
            return thAtoms[otrAtom.Index];
        }
    }

}
