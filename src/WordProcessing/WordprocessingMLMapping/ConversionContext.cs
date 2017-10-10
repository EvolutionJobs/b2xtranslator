using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.WordprocessingML;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class ConversionContext
    {
        private WordprocessingDocument _docx;
        //private Dictionary<Int32, SectionPropertyExceptions> _allSepx;
        //private Dictionary<Int32, ParagraphPropertyExceptions> _allPapx;
        private XmlWriterSettings _writerSettings;
        private WordDocument _doc;

        /// <summary>
        /// The source of the conversion.
        /// </summary>
        public WordDocument Doc
        {
            get { return _doc; }
            set { _doc = value; }
        }

        /// <summary>
        /// This is the target of the conversion.<br/>
        /// The result will be written to the parts of this document.
        /// </summary>
        public WordprocessingDocument Docx
        {
            get { return _docx; }
            set { _docx = value; }
        }

        /// <summary>
        /// The settings of the XmlWriter which writes to the part
        /// </summary>
        public XmlWriterSettings WriterSettings
        {
            get { return _writerSettings; }
            set { _writerSettings = value; }
        }

        /// <summary>
        /// A list thta contains all revision ids.
        /// </summary>
        public List<string> AllRsids;

        public ConversionContext(WordDocument doc)
        {
            this.Doc = doc;
            this.AllRsids = new List<string>();
        }

        /// <summary>
        /// Adds a new RSID to the list
        /// </summary>
        /// <param name="rsid"></param>
        public void AddRsid(string rsid)
        {
            if (!this.AllRsids.Contains(rsid))
                this.AllRsids.Add(rsid);
        }
    }
}
