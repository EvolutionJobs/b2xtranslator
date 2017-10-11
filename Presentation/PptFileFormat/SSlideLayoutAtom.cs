

using System;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    public enum SlideLayoutType
    {
        /// <summary>
        /// The slide is a title slide
        /// </summary>
        TitleSlide = 0,

        /// <summary>
        /// Title and body slide
        /// </summary>
        TitleAndBody = 1,

        /// <summary>
        /// Title master slide
        /// </summary>
        TitleMaster = 2,
        
        // 3 is unused

        /// <summary>
        /// Master notes layout
        /// </summary>
        MasterNotes = 4,

        /// <summary>
        /// Notes title/body layout
        /// </summary>
        NotesTitleAndBody = 5,

        /// <summary>
        /// Handout layout, therefore it doesn't have placeholders except header, footer, and date
        /// </summary>
        Handout = 6,

        /// <summary>
        /// Only title placeholder
        /// </summary>
        TitleOnly = 7,

        /// <summary>
        /// Body of the slide has 2 columns and a title
        /// </summary>
        TwoColumnsAndTitle = 8,

        /// <summary>
        /// Slide?s body has 2 rows and a title
        /// </summary>
        TwoRowsAndTitle = 9,

        /// <summary>
        /// Body contains 2 columns, right column has 2 rows
        /// </summary>
        TwoColumnsRightTwoRows = 10,

        /// <summary>
        /// Body contains 2 columns, left column has 2 rows
        /// </summary>
        TwoColumnsLeftTwoRows = 11,

        /// <summary>
        /// Body contains 2 rows, bottom row has 2 columns
        /// </summary>
        TwoRowsBottomTwoColumns = 12,

        /// <summary>
        /// Body contains 2 rows, top row has 2 columns
        /// </summary>
        TwoRowsTopTwoColumns = 13,

        /// <summary>
        /// 4 objects
        /// </summary>
        FourObjects = 14,

        /// <summary>
        /// Big object
        /// </summary>
        BigObject = 15,

        /// <summary>
        /// Blank slide
        /// </summary>
        Blank = 16,

        /// <summary>
        /// Vertical title on the right, body on the left
        /// </summary>
        VerticalTitleRightBodyLeft = 17,

        /// <summary>
        /// Vertical title on the right, body on the left split into 2 rows
        /// </summary>
        VerticalTitleRightBodyLeftTwoRows = 18
    }

    public class SSlideLayoutAtom
    {
        /// <summary>
        /// A SlideLayoutType that specifies a hint to the user interface which
        /// slide layout exists on the corresponding slide.
        /// 
        /// A slide layout specifies the type and number of placeholder shapes
        /// on a slide. A placeholder shape is specified as an OfficeArtSpContainer
        /// ([MS-ODRAW] section 2.2.14) that contains a PlaceholderAtom record
        /// with a pos field not equal to 0xFFFFFFFF. The placementId field of the
        /// PlaceholderAtom record specifies the placeholder shape type.
        /// </summary>
        public SlideLayoutType Geom;

        /// <summary>
        /// An array of PlaceholderEnum enumeration values that specifies
        /// a hint to the user interface which placeholder shapes exist on
        /// the corresponding slide.
        /// 
        /// The count of items in the array MUST be 8.
        /// </summary>
        public PlaceholderEnum[] PlaceholderTypes = new PlaceholderEnum[8];

        public SSlideLayoutAtom(BinaryReader reader)
        {
            int geom = reader.ReadInt32();
            this.Geom = (SlideLayoutType)geom;

            for (int i = 0; i < 8; i++)
                this.PlaceholderTypes[i] = (PlaceholderEnum)reader.ReadByte();
        }

        public override string ToString()
        {
            string s = string.Join(", ",
                Array.ConvertAll<PlaceholderEnum, string>(this.PlaceholderTypes,
                delegate(PlaceholderEnum pid) { return pid.ToString(); }));

            return string.Format("SSlideLayoutAtom(Geom = {0}, PlaceholderTypes = [{1}])",
                this.Geom, s);
        }
    }

}
