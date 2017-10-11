

using System.Collections.Generic;
using System.Xml;
using b2xtranslator.OpenXmlLib.PresentationML;
using b2xtranslator.PptFileFormat;

namespace b2xtranslator.PresentationMLMapping
{
    public class ConversionContext
    {
        private PresentationDocument _pptx;
        private XmlWriterSettings _writerSettings;
        private PowerpointDocument _ppt;

        private Dictionary<uint, MasterMapping> MasterIdToMapping = new Dictionary<uint, MasterMapping>();
        private Dictionary<uint, NotesMasterMapping> NotesMasterIdToMapping = new Dictionary<uint, NotesMasterMapping>();
        private Dictionary<uint, HandoutMasterMapping> HandoutMasterIdToMapping = new Dictionary<uint, HandoutMasterMapping>();
        public Dictionary<long, string> AddedImages = new Dictionary<long, string>();
        public int lastImageID = 1000;

        /// <summary>
        /// The source of the conversion.
        /// </summary>
        public PowerpointDocument Ppt
        {
            get { return this._ppt; }
            set { this._ppt = value; }
        }

        /// <summary>
        /// This is the target of the conversion.<br/>
        /// The result will be written to the parts of this document.
        /// </summary>
        public PresentationDocument Pptx
        {
            get { return this._pptx; }
            set { this._pptx = value; }
        }

        /// <summary>
        /// The settings of the XmlWriter which writes to the part
        /// </summary>
        public XmlWriterSettings WriterSettings
        {
            get { return this._writerSettings; }
            set { this._writerSettings = value; }
        }

        public ConversionContext(PowerpointDocument ppt)
        {
            this.Ppt = ppt;
        }

        /// <summary>
        /// Registers a NotesMasterMapping so it can later be looked up by its master ID.
        /// </summary>
        /// <param name="masterId">Master id with which to associate the MasterMapping.</param>
        /// <param name="mapping">MasterMapping to be registered.</param>
        public void RegisterNotesMasterMapping(uint masterId, NotesMasterMapping mapping)
        {
            this.NotesMasterIdToMapping[masterId] = mapping;
        }

        /// <summary>
        /// Registers a HandoutMasterMapping so it can later be looked up by its master ID.
        /// </summary>
        /// <param name="masterId">Master id with which to associate the MasterMapping.</param>
        /// <param name="mapping">MasterMapping to be registered.</param>
        public void RegisterHandoutMasterMapping(uint masterId, HandoutMasterMapping mapping)
        {
            this.HandoutMasterIdToMapping[masterId] = mapping;
        }

        /// <summary>
        /// Returns the NotesMasterMapping associated with the specified master ID.
        /// </summary>
        /// <param name="masterId">Master ID for which to find a MasterMapping.</param>
        /// <returns>Found MasterMapping or null if none was found.</returns>
        public NotesMasterMapping GetNotesMasterMappingByMasterId(uint masterId)
        {
            return this.NotesMasterIdToMapping[masterId];
        }

        /// <summary>
        /// Returns the HandoutMasterMapping associated with the specified master ID.
        /// </summary>
        /// <param name="masterId">Master ID for which to find a MasterMapping.</param>
        /// <returns>Found MasterMapping or null if none was found.</returns>
        public HandoutMasterMapping GetHandoutMasterMappingByMasterId(uint masterId)
        {
            return this.HandoutMasterIdToMapping[masterId];
        }

        /// <summary>
        /// Registers a MasterMapping so it can later be looked up by its master ID.
        /// </summary>
        /// <param name="masterId">Master id with which to associate the MasterMapping.</param>
        /// <param name="mapping">MasterMapping to be registered.</param>
        public void RegisterMasterMapping(uint masterId, MasterMapping mapping)
        {
            this.MasterIdToMapping[masterId] = mapping;
        }

        /// <summary>
        /// Returns the MasterMapping associated with the specified master ID.
        /// </summary>
        /// <param name="masterId">Master ID for which to find a MasterMapping.</param>
        /// <returns>Found MasterMapping or null if none was found.</returns>
        public MasterMapping GetMasterMappingByMasterId(uint masterId)
        {
            return this.MasterIdToMapping[masterId];
        }

        /// <summary>
        /// Returns the MasterMapping associated with the specified master ID if it exists.
        /// Else a new MasterMapping is created.
        /// </summary>
        /// <param name="masterId">Master ID for which to find or create a MasterMapping.</param>
        /// <returns>Found or created MasterMapping.</returns>
        public MasterMapping GetOrCreateMasterMappingByMasterId(uint masterId)
        {
            if (!this.MasterIdToMapping.ContainsKey(masterId))
                this.MasterIdToMapping[masterId] = new MasterMapping(this);

            return this.MasterIdToMapping[masterId];
        }

        /// <summary>
        /// Returns the NotesMasterMapping associated with the specified master ID if it exists.
        /// Else a new NotesMasterMapping is created.
        /// </summary>
        /// <param name="masterId">Master ID for which to find or create a NotesMasterMapping.</param>
        /// <returns>Found or created NotesMasterMapping.</returns>
        public NotesMasterMapping GetOrCreateNotesMasterMappingByMasterId(uint masterId)
        {
            if (!this.NotesMasterIdToMapping.ContainsKey(masterId))
                this.NotesMasterIdToMapping[masterId] = new NotesMasterMapping(this);

            return this.NotesMasterIdToMapping[masterId];
        }

        /// <summary>
        /// Returns the HandoutMasterMapping associated with the specified master ID if it exists.
        /// Else a new HandoutMasterMapping is created.
        /// </summary>
        /// <param name="masterId">Master ID for which to find or create a HandoutMasterMapping.</param>
        /// <returns>Found or created HandoutMasterMapping.</returns>
        public HandoutMasterMapping GetOrCreateHandoutMasterMappingByMasterId(uint masterId)
        {
            if (!this.HandoutMasterIdToMapping.ContainsKey(masterId))
                this.HandoutMasterIdToMapping[masterId] = new HandoutMasterMapping(this);

            return this.HandoutMasterIdToMapping[masterId];
        }

        protected Dictionary<uint, MasterLayoutManager> MasterIdToLayoutManager =
            new Dictionary<uint, MasterLayoutManager>();

        public MasterLayoutManager GetOrCreateLayoutManagerByMasterId(uint masterId)
        {
            if (!this.MasterIdToLayoutManager.ContainsKey(masterId))
                this.MasterIdToLayoutManager[masterId] = new MasterLayoutManager(this, masterId);

            return this.MasterIdToLayoutManager[masterId];
        }
    }

    public class MasterLayoutManager
    {
        protected ConversionContext _ctx;
        protected uint MasterId;

        /// <summary>
        /// PPT2007 layouts are stored inline with the master and
        /// have an instance id for associating them with slides.
        /// </summary>
        public Dictionary<uint, SlideLayoutPart> InstanceIdToLayoutPart =
            new Dictionary<uint, SlideLayoutPart>();

        /// <summary>
        /// Pre-PPT2007 layouts are specified in SSlideLayoutAtom
        /// as a SlideLayoutType integer value. Each SlideLayoutType
        /// can be mapped to a layout XML file together with a list of
        /// placeholder types.
        /// 
        /// This dictionary is used for associating default layout
        /// part filenames with layout parts.
        /// </summary>
        public Dictionary<string, SlideLayoutPart> LayoutFilenameToLayoutPart =
            new Dictionary<string, SlideLayoutPart>();

        /// <summary>
        /// Pre-PPT2007 TitleMaster slides need to be converted to
        /// SlideLayoutParts for OOXML. The SlideLayoutParts for
        /// TitleMaster slides are stored in this dictionary.
        /// </summary>
        public Dictionary<uint, SlideLayoutPart> TitleMasterIdToLayoutPart =
            new Dictionary<uint, SlideLayoutPart>();

        public Dictionary<string, SlideLayoutPart> CodeToLayoutPart =
            new Dictionary<string, SlideLayoutPart>();

        public MasterLayoutManager(ConversionContext ctx, uint masterId)
        {
            this._ctx = ctx;
            this.MasterId = masterId;
        }

        public List<SlideLayoutPart> GetAllLayoutParts()
        {
            var result = new List<SlideLayoutPart>();

            result.AddRange(this.InstanceIdToLayoutPart.Values);
            result.AddRange(this.LayoutFilenameToLayoutPart.Values);
            result.AddRange(this.TitleMasterIdToLayoutPart.Values);
            result.AddRange(this.CodeToLayoutPart.Values);

            return result;
        }

        public SlideLayoutPart AddLayoutPartWithInstanceId(uint instanceId)
        {
            var masterPart = this._ctx.GetOrCreateMasterMappingByMasterId(this.MasterId).MasterPart;
            var layoutPart = masterPart.AddSlideLayoutPart();

            this.InstanceIdToLayoutPart.Add(instanceId, layoutPart);
            return layoutPart;
        }

        public SlideLayoutPart GetLayoutPartByInstanceId(uint instanceId)
        {
            return this.InstanceIdToLayoutPart[instanceId];
        }

        public SlideLayoutPart GetOrCreateLayoutPartByLayoutType(SlideLayoutType type,
            PlaceholderEnum[] placeholderTypes)
        {
            var masterPart = this._ctx.GetOrCreateMasterMappingByMasterId(this.MasterId).MasterPart;
            string layoutFilename = Utils.SlideLayoutTypeToFilename(type, placeholderTypes);

            
            if (!this.LayoutFilenameToLayoutPart.ContainsKey(layoutFilename))
            {
                var slideLayoutDoc = Utils.GetDefaultDocument("slideLayouts." + layoutFilename);

                var layoutPart = masterPart.AddSlideLayoutPart();
                slideLayoutDoc.WriteTo(layoutPart.XmlWriter);
                layoutPart.XmlWriter.Flush();

                this.LayoutFilenameToLayoutPart.Add(layoutFilename, layoutPart);
            }

            return this.LayoutFilenameToLayoutPart[layoutFilename];
        }

        public SlideLayoutPart GetOrCreateLayoutPartByCode(string xml)
        {
            var masterPart = this._ctx.GetOrCreateMasterMappingByMasterId(this.MasterId).MasterPart;

            if (!this.CodeToLayoutPart.ContainsKey(xml))
            {

                var doc = new XmlDocument();
                doc.LoadXml(xml);

                var layoutPart = masterPart.AddSlideLayoutPart();
                doc.WriteTo(layoutPart.XmlWriter);
                layoutPart.XmlWriter.Flush();
                this.CodeToLayoutPart.Add(xml, layoutPart);
            }

            return this.CodeToLayoutPart[xml];
        }

        public SlideLayoutPart GetOrCreateLayoutPartForTitleMasterId(uint titleMasterId)
        {
            var masterPart = this._ctx.GetOrCreateMasterMappingByMasterId(this.MasterId).MasterPart;

            if (!this.TitleMasterIdToLayoutPart.ContainsKey(titleMasterId))
            {
                var titleMaster = this._ctx.Ppt.FindMasterRecordById(titleMasterId);
                var layoutPart = masterPart.AddSlideLayoutPart();
                new TitleMasterMapping(this._ctx, layoutPart).Apply(titleMaster);
                this.TitleMasterIdToLayoutPart.Add(titleMasterId, layoutPart);
            }

            return this.TitleMasterIdToLayoutPart[titleMasterId];
        }
    }
}
