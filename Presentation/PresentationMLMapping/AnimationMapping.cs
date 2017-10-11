

using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OfficeDrawing;
using System.Drawing;

namespace b2xtranslator.PresentationMLMapping
{
    class AnimationMapping :
        AbstractOpenXmlMapping//,
    {
        protected ConversionContext _ctx;
        private ShapeTreeMapping _stm;
        private List<Point> TextAreasForAnimation = new List<Point>();

        public AnimationMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            this._ctx = ctx;
        }

        public void Apply(ProgBinaryTagDataBlob blob, PresentationMapping<RegularContainer> parentMapping, Dictionary<AnimationInfoContainer, int> animations, ShapeTreeMapping stm)
        {
            this._parentMapping = parentMapping;
            this._stm = stm;
            var animAtoms = new Dictionary<AnimationInfoAtom, int>();
            foreach (var container in animations.Keys)
            {
                var anim = container.FirstChildWithType<AnimationInfoAtom>();
                animAtoms.Add(anim, animations[container]);
            }

             var c1 = blob.FirstChildWithType<ExtTimeNodeContainer>();
             if (c1 != null)
             {
                 var c2 = c1.FirstChildWithType<ExtTimeNodeContainer>();
                 if (c2 != null)
                 {
                     var c3 = c2.FirstChildWithType<ExtTimeNodeContainer>();
                     if (c3 != null)
                     {
                         writeTiming(animAtoms, blob);
                     }
                 }
             }
        }

        private PresentationMapping<RegularContainer> _parentMapping;

        private int lastID = 0;
        private void writeTiming(Dictionary<AnimationInfoAtom, int> blindAtoms, ProgBinaryTagDataBlob blob)
        {
            this.lastID = 0;

            this._writer.WriteStartElement("p", "timing", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "tnLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());
            this._writer.WriteAttributeString("dur", "indefinite");
            //_writer.WriteAttributeString("restart", "never");
            this._writer.WriteAttributeString("nodeType", "tmRoot");

            this._writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "seq", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("concurrent", "1");
            this._writer.WriteAttributeString("nextAc", "seek");

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());
            this._writer.WriteAttributeString("dur", "indefinite");
            this._writer.WriteAttributeString("nodeType", "mainSeq");

            this._writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

            if (blob != null)
            {

                var c1 = blob.FirstChildWithType<ExtTimeNodeContainer>();
                if (c1 != null)
                {
                    var c2 = c1.FirstChildWithType<ExtTimeNodeContainer>();
                    if (c2 != null)
                    {

                        foreach (var c3 in c2.AllChildrenWithType<ExtTimeNodeContainer>())
                        if (c3 != null)
                        {

                            int counter = 0;
                            AnimationInfoAtom a;
                            var atoms = new List<AnimationInfoAtom>();
                            foreach (var key in blindAtoms.Keys) atoms.Add(key);

                                this._writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

                                this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
                                this._writer.WriteAttributeString("id", (++this.lastID).ToString());
                                this._writer.WriteAttributeString("fill", "hold");

                                this._writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);


                            foreach (var c in c3.AllChildrenWithType<TimeConditionContainer>())
                            {
                                var t = c.FirstChildWithType<TimeConditionAtom>();

                                    this._writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

                                switch (t.triggerEvent)
                                {
                                    case 0x0: //none
                                        break;
                                    case 0x1: //onBegin
                                            this._writer.WriteAttributeString("evt", "onBegin");
                                        break;
                                    case 0x3: //Start
                                            this._writer.WriteAttributeString("evt", "begin");
                                        break;
                                    case 0x4: //End
                                            this._writer.WriteAttributeString("evt", "end");
                                        break;
                                    case 0x5: //Mouse click
                                            this._writer.WriteAttributeString("evt", "onClick");
                                        break;
                                    case 0x7: //Mouse over
                                            this._writer.WriteAttributeString("evt", "onMouseOver");
                                        break;
                                    case 0x9: //OnNext
                                            this._writer.WriteAttributeString("evt", "onNext");
                                        break;
                                    case 0xa: //OnPrev
                                            this._writer.WriteAttributeString("evt", "onPrev");
                                        break;
                                    case 0xb: //Stop audio
                                            this._writer.WriteAttributeString("evt", "onStopAudio");
                                        break;
                                    default:
                                        break;
                                }

                                if (t.delay == -1)
                                {
                                        this._writer.WriteAttributeString("delay", "indefinite");
                                }
                                else
                                {
                                        this._writer.WriteAttributeString("delay", t.delay.ToString());
                                }

                                if (t.triggerObject == TimeConditionAtom.TriggerObjectEnum.TimeNode)
                                {
                                        this._writer.WriteStartElement("p", "tn", OpenXmlNamespaces.PresentationML);
                                        this._writer.WriteAttributeString("val", t.id.ToString());
                                        this._writer.WriteEndElement();
                                }

                                    this._writer.WriteEndElement(); //cond

                            }

                                this._writer.WriteEndElement(); //stCondLst

                                this._writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);


                            foreach (var c4 in c3.AllChildrenWithType<ExtTimeNodeContainer>())
                            {
                                a = null;
                                if (atoms.Count > counter) a = atoms[counter];
                                writePar(c4, a);
                                counter++;
                            }

                                this._writer.WriteEndElement(); //childTnLst

                                this._writer.WriteEndElement(); //cTn

                                this._writer.WriteEndElement(); //par
                        }
                    }
                }
            }

            this._writer.WriteEndElement(); //childTnLst

            this._writer.WriteEndElement(); //cTn

            this._writer.WriteStartElement("p", "prevCondLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("evt", "onPrev");
            this._writer.WriteAttributeString("delay", "0");

            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            this._writer.WriteElementString("p", "sldTgt", OpenXmlNamespaces.PresentationML, "");

            this._writer.WriteEndElement(); //tgtEl

            this._writer.WriteEndElement(); //cond

            this._writer.WriteEndElement(); //prevCondLst

            this._writer.WriteStartElement("p", "nextCondLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("evt", "onNext");
            this._writer.WriteAttributeString("delay", "0");

            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            this._writer.WriteElementString("p", "sldTgt", OpenXmlNamespaces.PresentationML, "");

            this._writer.WriteEndElement(); //tgtEl

            this._writer.WriteEndElement(); //cond

            this._writer.WriteEndElement(); //nextCondLst

            this._writer.WriteEndElement(); //seq

            this._writer.WriteEndElement(); //childTnLst

            this._writer.WriteEndElement(); //cTn

            this._writer.WriteEndElement(); //par

            this._writer.WriteEndElement(); //tnLst

            if (blindAtoms.Count > 0)
            {

                this._writer.WriteStartElement("p", "bldLst", OpenXmlNamespaces.PresentationML);

                foreach (var animinfo in blindAtoms.Keys)
                {
                    this._writer.WriteStartElement("p", "bldP", OpenXmlNamespaces.PresentationML);
                    this._writer.WriteAttributeString("spid", blindAtoms[animinfo].ToString());
                    this._writer.WriteAttributeString("grpId", "0");

                    if (animinfo.animBuildType == AnimationInfoAtom.AnimBuildTypeEnum.Level1Build) this._writer.WriteAttributeString("build", "p");
                    if (animinfo.fAnimateBg) this._writer.WriteAttributeString("animBg", "1");

                    this._writer.WriteEndElement(); //bldP
                }

                this._writer.WriteEndElement(); //bldLst

            }

            this._writer.WriteEndElement(); //timing
        }


        private void writePar(ExtTimeNodeContainer container, AnimationInfoAtom animinfo)
        {

            this._writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());
            this._writer.WriteAttributeString("fill", "hold");

            this._writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            if (container.FirstChildWithType<TimeConditionContainer>() != null)
            {
                this._writer.WriteAttributeString("delay", container.FirstChildWithType<TimeConditionContainer>().FirstChildWithType<TimeConditionAtom>().delay.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("delay", "0");
            }

            this._writer.WriteEndElement(); //cond

            this._writer.WriteEndElement(); //stCondLst

            this._writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

            foreach (var c2 in container.AllChildrenWithType<ExtTimeNodeContainer>())
            {
                var Attributes = new Dictionary<TimePropertyID4TimeNode, TimeVariantValue>();
                foreach (var tv in c2.FirstChildWithType<TimePropertyList4TimeNodeContainer>().AllChildrenWithType<TimeVariantValue>())
                {
                    Attributes.Add((TimePropertyID4TimeNode)tv.Instance, tv);
                }

                this._writer.WriteStartElement("p", "par", OpenXmlNamespaces.PresentationML);

                this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
                this._writer.WriteAttributeString("id", (++this.lastID).ToString());

                string filter = "";
                foreach (var c3 in c2.AllChildrenWithType<ExtTimeNodeContainer>())
                {
                    foreach (var c4 in c3.AllChildrenWithType<TimeEffectBehaviorContainer>())
                    {
                        foreach (var v in c4.AllChildrenWithType<TimeVariantValue>())
                        {
                            if (v.type == TimeVariantTypeEnum.String)
                            {
                                filter = v.stringValue;
                            }
                        }
                    }
                }


                if (Attributes[TimePropertyID4TimeNode.EffectID] != null)
                {
                    this._writer.WriteAttributeString("presetID", (Attributes[TimePropertyID4TimeNode.EffectID].intValue).ToString()); //3
                }
                else
                {
                    this._writer.WriteAttributeString("presetID", "12"); //3
                }

                switch (Attributes[TimePropertyID4TimeNode.EffectType].intValue)
                {
                    case 1:
                        this._writer.WriteAttributeString("presetClass", "entr");
                        break;
                    case 2:
                        this._writer.WriteAttributeString("presetClass", "exit");
                        break;
                    case 3:
                        this._writer.WriteAttributeString("presetClass", "emph");
                        break;
                    case 4:
                        this._writer.WriteAttributeString("presetClass", "path");
                        break;
                    case 5:
                        this._writer.WriteAttributeString("presetClass", "verb");
                        break;
                    case 6:
                        this._writer.WriteAttributeString("presetClass", "mediacall");
                        break;
                }

                if (Attributes.ContainsKey(TimePropertyID4TimeNode.EffectDir) && Attributes[TimePropertyID4TimeNode.EffectDir] != null)
                {
                    this._writer.WriteAttributeString("presetSubtype", (Attributes[TimePropertyID4TimeNode.EffectDir].intValue).ToString()); 
                }
                else
                {
                    this._writer.WriteAttributeString("presetSubtype", "4");
                }


                this._writer.WriteAttributeString("fill", "hold");
                //_writer.WriteAttributeString("grpId", "0");

                bool nodeTypeWritten = false;
                if (container.FirstChildWithType<ExtTimeNodeContainer>() != null)
                {
                    if (c2.FirstChildWithType<TimePropertyList4TimeNodeContainer>() != null)
                    {
                                                                      
                        switch (Attributes[TimePropertyID4TimeNode.EffectNodeType].intValue)
                        {
                            case 1:
                                this._writer.WriteAttributeString("nodeType", "clickEffect");
                                nodeTypeWritten = true;
                                break;
                            case 2:
                                nodeTypeWritten = true;
                                this._writer.WriteAttributeString("nodeType", "withEffect");
                                break;
                            case 3:
                                this._writer.WriteAttributeString("nodeType", "afterEffect");
                                nodeTypeWritten = true;
                                break;
                            case 4:
                            case 5:
                            case 6:
                            case 7:
                            case 8:
                            case 9:
                            default:
                                break;
                        }

                    }

                }
                if (!nodeTypeWritten) this._writer.WriteAttributeString("nodeType", "clickEffect");

                this._writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

                this._writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

                this._writer.WriteAttributeString("delay", "0");

                this._writer.WriteEndElement(); //cond

                this._writer.WriteEndElement(); //stCondLst


                if (c2.FirstChildWithType<TimeIterateDataAtom>() != null)
                {
                    var tida = c2.FirstChildWithType<TimeIterateDataAtom>();

                    this._writer.WriteStartElement("p", "iterate", OpenXmlNamespaces.PresentationML);

                    if (tida.fIterateTypePropertyUsed)
                    {
                        switch (tida.iterateType)
                        {
                            case 0:
                                this._writer.WriteAttributeString("type", "el");
                                break;
                            case 1:
                                this._writer.WriteAttributeString("type", "wd");
                                break;
                            case 2:
                                this._writer.WriteAttributeString("type", "lt");
                                break;
                        }
                    }

                    this._writer.WriteStartElement("p", "tmPct", OpenXmlNamespaces.PresentationML);

                    this._writer.WriteAttributeString("val", (tida.iterateInterval * 1000).ToString());
                    this._writer.WriteEndElement(); //tmPct
                    this._writer.WriteEndElement(); //iterate
                }


                this._writer.WriteStartElement("p", "childTnLst", OpenXmlNamespaces.PresentationML);

                int targetRun = -1;

                if (c2.FirstChildWithType<ExtTimeNodeContainer>().FirstChildWithType<TimeRotationBehaviorContainer>() != null)
                {
                    writeAnimRot(c2, ref targetRun, c2.FirstChildWithType<ExtTimeNodeContainer>().FirstChildWithType<TimeRotationBehaviorContainer>().FirstChildWithType<TimeRotationBehaviorAtom>());
                }

                if (c2.FirstChildWithType<ExtTimeNodeContainer>().FirstChildWithType<TimeCommandBehaviorContainer>() != null)
                {
                    writeAnimCmd(c2, ref targetRun, c2.FirstChildWithType<ExtTimeNodeContainer>().FirstChildWithType<TimeCommandBehaviorContainer>().FirstChildWithType<TimeCommandBehaviorAtom>());
                }
               
                writeAnimations(c2, targetRun);

                this._writer.WriteEndElement(); //childTnLst

                //slaves
                if (c2.AllChildrenWithType<SlaveContainer>().Count > 0)
                {
                    if (c2.FirstDescendantWithType<TimeColorBehaviorContainer>() != null)
                    {

                        this._writer.WriteStartElement("p", "subTnLst", OpenXmlNamespaces.PresentationML);
                        foreach (var sc in c2.AllChildrenWithType<SlaveContainer>())
                        {                           

                            var tcbc = sc.FirstChildWithType<TimeColorBehaviorContainer>();
                            if (tcbc != null)
                            {
                                writeColor(sc, targetRun);                                                                
                            }

                        }
                        this._writer.WriteEndElement(); //subTnLst
                    }
                }


                this._writer.WriteEndElement(); //cTn

                this._writer.WriteEndElement(); //par

            }

            this._writer.WriteEndElement(); //childTnLst

            this._writer.WriteEndElement(); //cTn

            this._writer.WriteEndElement(); //par

       }

        private string getShapeId(uint c4Id)
        {
            if (this._stm.spidToId.ContainsKey((int)c4Id))
            {
                return this._stm.spidToId[(int)c4Id].ToString();
            }
            else
            {
                foreach (int sId in this._stm.spidToId.Keys)
                {
                    if (sId > 0)
                    {
                        return this._stm.spidToId[sId].ToString();
                    }
                }
            }
            return "";
        }

        private void writeSet(ExtTimeNodeContainer c, ref int targetRun)
        {           

            var tna = c.FirstChildWithType<TimeNodeAtom>();

            var tsbc = c.FirstChildWithType<TimeSetBehaviourContainer>();
            var val = tsbc.FirstChildWithType<TimeVariantValue>();
            var tbc = tsbc.FirstChildWithType<TimeBehaviorContainer>();
            var tba = tbc.FirstChildWithType<TimeBehaviorAtom>();
            var attrName = tbc.FirstChildWithType<TimeStringListContainer>().FirstChildWithType<TimeVariantValue>();
            var vsa = tbc.FirstChildWithType<ClientVisualElementContainer>().FirstChildWithType<VisualShapeAtom>();

            TimeConditionAtom tca = null;
            if (c.FirstChildWithType<TimeConditionContainer>() != null)
            tca = c.FirstChildWithType<TimeConditionContainer>().FirstChildWithType<TimeConditionAtom>();

            this._writer.WriteStartElement("p", "set", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);

            if (tba.fAdditivePropertyUsed)
            {
                switch (tba.behaviorAdditive)
                {
                    case 0: //override
                        this._writer.WriteAttributeString("additive", "base");
                        break;
                    case 1: //add
                        this._writer.WriteAttributeString("additive", "sum");
                        break;
                }
            }

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());

            if (tna.fDurationProperty)
            {
                this._writer.WriteAttributeString("dur", tna.duration.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("dur", "1");
            }


            if (tna.fFillProperty)
            {
                switch (tna.fill)
                {
                    case 0:
                    case 3:
                        this._writer.WriteAttributeString("fill", "hold");
                        break;
                    case 1:
                    case 4:
                        this._writer.WriteAttributeString("fill", "reset");
                        break;
                    case 2:
                        this._writer.WriteAttributeString("fill", "freeze"); //TODO:verify
                        break;
                }
            }
            else
            {
                this._writer.WriteAttributeString("fill", "hold");
            }

            this._writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            if (tca == null)
            {
                this._writer.WriteAttributeString("delay", "0"); 
            }
            else
            {
                this._writer.WriteAttributeString("delay", tca.delay.ToString());
            }

            this._writer.WriteEndElement(); //cond

            this._writer.WriteEndElement(); //stCondLst

            this._writer.WriteEndElement(); //cTn

            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("spid", getShapeId(c.FirstDescendantWithType<VisualShapeAtom>().shapeIdRef));

            CheckAndWriteStartEndRuns(c, ref targetRun);

            this._writer.WriteEndElement(); //spTgt

            this._writer.WriteEndElement(); //tgtEl

            this._writer.WriteStartElement("p", "attrNameLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteElementString("p", "attrName", OpenXmlNamespaces.PresentationML, attrName.stringValue);

            this._writer.WriteEndElement(); //attrNameLst

            this._writer.WriteEndElement(); //cBhvr

            this._writer.WriteStartElement("p", "to", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "strVal", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("val", val.stringValue);

            this._writer.WriteEndElement(); //str

            this._writer.WriteEndElement(); //to

            this._writer.WriteEndElement(); //set
        }

        private void writeAnimRot(ExtTimeNodeContainer c, ref int targetRun, TimeRotationBehaviorAtom trba)
        {
            var tna = c.FirstChildWithType<TimeNodeAtom>();

            this._writer.WriteStartElement("p", "animRot", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("by", (trba.fBy * 60000).ToString("#"));

            this._writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());
            
            if (tna.fDurationProperty)
            {
                this._writer.WriteAttributeString("dur", tna.duration.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("dur", "500");
            }

            this._writer.WriteAttributeString("fill", "hold");
            this._writer.WriteEndElement(); //cTn

            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("spid", getShapeId(c.FirstDescendantWithType<VisualShapeAtom>().shapeIdRef));

            CheckAndWriteStartEndRuns(c, ref targetRun);

            this._writer.WriteEndElement(); //spTgt

            this._writer.WriteEndElement(); //tgtEl

            this._writer.WriteStartElement("p", "attrNameLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteElementString("p", "attrName", OpenXmlNamespaces.PresentationML, "r");

            this._writer.WriteEndElement(); //attrNameLst

            this._writer.WriteEndElement(); //cBhvr

            this._writer.WriteEndElement(); //animRot
        }

        private void writeAnimCmd(ExtTimeNodeContainer c, ref int targetRun, TimeCommandBehaviorAtom tcba)
        {
            var tna = c.FirstChildWithType<TimeNodeAtom>();

            this._writer.WriteStartElement("p", "cmd", OpenXmlNamespaces.PresentationML);

            if (tcba.fCommandPropertyUsed)
            {
                switch (tcba.commandBehaviorType)
                {
                    case 0:
                        this._writer.WriteAttributeString("type", "evt");
                        break;
                    case 1:
                        this._writer.WriteAttributeString("type", "call");
                        break;
                    case 2:
                        this._writer.WriteAttributeString("type", "verb");
                        break;
                }
            }
            

            if (tcba.fCommandPropertyUsed)
            {
                this._writer.WriteAttributeString("cmd", tcba.FirstAncestorWithType<TimeCommandBehaviorContainer>().FirstChildWithType<TimeVariantValue>().stringValue);
            }

            this._writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());

            if (tna.fDurationProperty)
            {
                this._writer.WriteAttributeString("dur", tna.duration.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("dur", "1");
            }

            this._writer.WriteAttributeString("fill", "hold");
            this._writer.WriteEndElement(); //cTn

            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("spid", getShapeId(c.FirstDescendantWithType<VisualShapeAtom>().shapeIdRef));

            CheckAndWriteStartEndRuns(c, ref targetRun);

            this._writer.WriteEndElement(); //spTgt

            this._writer.WriteEndElement(); //tgtEl

            this._writer.WriteEndElement(); //cBhvr

            this._writer.WriteEndElement(); //animRot
        }

        public void writeAnim(ExtTimeNodeContainer c, int targetRun)
        {
            var tna = c.FirstChildWithType<TimeNodeAtom>();
            var bc = c.FirstChildWithType<TimeAnimateBehaviourContainer>();

            var tbc = bc.FirstChildWithType<TimeBehaviorContainer>();
            var tba = tbc.FirstChildWithType<TimeBehaviorAtom>();
            var taba = bc.FirstChildWithType<TimeAnimateBehaviorAtom>();
            var attrName = tbc.FirstChildWithType<TimeStringListContainer>().FirstChildWithType<TimeVariantValue>();

            string filter = "";
            if (c.FirstChildWithType<TimePropertyList4TimeNodeContainer>() != null && c.FirstChildWithType<TimePropertyList4TimeNodeContainer>().FirstChildWithType<TimeVariantValue>() != null)
            {
                filter = c.FirstChildWithType<TimePropertyList4TimeNodeContainer>().FirstChildWithType<TimeVariantValue>().stringValue;
            }
            var lst = new List<Record>();
            string fieldName = "";
            if (bc.FirstChildWithType<TimeAnimationValueListContainer>() != null)
            {
                lst = bc.FirstChildWithType<TimeAnimationValueListContainer>().Children;
            }
            else
            {
                fieldName = tbc.FirstChildWithType<TimeStringListContainer>().FirstChildWithType<TimeVariantValue>().stringValue;
            }

            this._writer.WriteStartElement("p", "anim", OpenXmlNamespaces.PresentationML);

            if (taba.fToPropertyUsed)
            {
                this._writer.WriteAttributeString("to", bc.FirstChildWithType<TimeVariantValue>().stringValue);
            }

            if (taba.fCalcModePropertyUsed)
            {
                switch (taba.calcMode)
                {
                    case 0: //discrete
                        this._writer.WriteAttributeString("calcmode", "discrete");
                        break;
                    case 1: //linear
                        this._writer.WriteAttributeString("calcmode", "lin");
                        break;
                    case 2: //formula
                        this._writer.WriteAttributeString("calcmode", "fmla");
                        break;
                }
            }
            else
            {
                //default
                this._writer.WriteAttributeString("calcmode", "lin");
            }

            if (taba.fValueTypePropertyUsed)
            {
                switch (taba.valueType)
                {
                    case TimeAnimateBehaviorValueTypeEnum.Color:
                        this._writer.WriteAttributeString("valueType", "clr");
                        break;
                    case TimeAnimateBehaviorValueTypeEnum.Number:
                        this._writer.WriteAttributeString("valueType", "num");
                        break;
                    case TimeAnimateBehaviorValueTypeEnum.String:
                        this._writer.WriteAttributeString("valueType", "str");
                        break;
                }
            }
            else
            {
                this._writer.WriteAttributeString("valueType", "num");
            }

            this._writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);

            if (tba.fAdditivePropertyUsed)
            {
                switch (tba.behaviorAdditive)
                {
                    case 0: //override
                        this._writer.WriteAttributeString("additive", "base");
                        break;
                    case 1: //add
                        this._writer.WriteAttributeString("additive", "sum");
                        break;
                }
            }

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());

            if (tna.fDurationProperty)
            {
                this._writer.WriteAttributeString("dur", tna.duration.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("dur", "500");
            }

            if (filter.Length > 0)
            {
                this._writer.WriteAttributeString("tmFilter", filter);
            }

            if (tna.fFillProperty)
            {
                switch (tna.fill)
                {
                    case 0:
                    case 3:
                        this._writer.WriteAttributeString("fill", "hold");
                        break;
                    case 1:
                    case 4:
                        this._writer.WriteAttributeString("fill", "reset");
                        break;
                    case 2:
                        this._writer.WriteAttributeString("fill", "freeze"); //TODO:verify
                        break;
                }
            }
            else
            {
                this._writer.WriteAttributeString("fill", "hold");
            }

            this._writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            if (c.FirstChildWithType<TimeConditionContainer>() != null)
            {
                this._writer.WriteAttributeString("delay", c.FirstChildWithType<TimeConditionContainer>().FirstChildWithType<TimeConditionAtom>().delay.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("delay", "0");
            }

            this._writer.WriteEndElement(); //cond

            this._writer.WriteEndElement(); //stCondLst


            this._writer.WriteEndElement(); //cTn

            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("spid", getShapeId(c.FirstDescendantWithType<VisualShapeAtom>().shapeIdRef));
            CheckAndWriteStartEndRuns(c, ref targetRun);

            this._writer.WriteEndElement(); //spTgt
            this._writer.WriteEndElement(); //tgtEl

            this._writer.WriteStartElement("p", "attrNameLst", OpenXmlNamespaces.PresentationML);
            this._writer.WriteElementString("p", "attrName", OpenXmlNamespaces.PresentationML, attrName.stringValue);

            this._writer.WriteEndElement(); //attrNameLst

            this._writer.WriteEndElement(); //cBhvr

            if (lst.Count > 0)
            {
                this._writer.WriteStartElement("p", "tavLst", OpenXmlNamespaces.PresentationML);

                TimeAnimationValueAtom tava;
                TimeVariantValue tvv;
                TimeVariantValue tvv2 = null;
                while (lst.Count > 0)
                {
                    tava = (TimeAnimationValueAtom)lst[0];
                    tvv = (TimeVariantValue)lst[1];
                    if (lst.Count > 2 && lst[2] is TimeVariantValue)
                    {
                        tvv2 = (TimeVariantValue)lst[2];
                        lst.RemoveAt(0);
                    } else {
                        tvv2 = null;
                    }
                    lst.RemoveAt(0);
                    lst.RemoveAt(0);

                    this._writer.WriteStartElement("p", "tav", OpenXmlNamespaces.PresentationML);
                    this._writer.WriteAttributeString("tm", (tava.time * 100).ToString());

                    if (tvv2 != null && tvv2.type == TimeVariantTypeEnum.String && tvv2.stringValue.Length > 0)
                    {
                        this._writer.WriteAttributeString("fmla", tvv2.stringValue);
                    }

                    this._writer.WriteStartElement("p", "val", OpenXmlNamespaces.PresentationML);

                    switch (tvv.type)
                    {
                        case TimeVariantTypeEnum.Bool:
                            this._writer.WriteStartElement("p", "boolVal", OpenXmlNamespaces.PresentationML);
                            this._writer.WriteAttributeString("val", tvv.boolValue.ToString());
                            break;
                        case TimeVariantTypeEnum.Float:
                            this._writer.WriteStartElement("p", "fltVal", OpenXmlNamespaces.PresentationML);
                            this._writer.WriteAttributeString("val", tvv.floatValue.ToString());
                            break;
                        case TimeVariantTypeEnum.Int:
                            this._writer.WriteStartElement("p", "intVal", OpenXmlNamespaces.PresentationML);
                            this._writer.WriteAttributeString("val", tvv.intValue.ToString());
                            break;
                        case TimeVariantTypeEnum.String:
                            this._writer.WriteStartElement("p", "strVal", OpenXmlNamespaces.PresentationML);
                            this._writer.WriteAttributeString("val", tvv.stringValue);
                            break;
                    }

                    this._writer.WriteEndElement(); //strVal
                    this._writer.WriteEndElement(); //val
                    this._writer.WriteEndElement(); //tav
                }

                this._writer.WriteEndElement(); //tavLst
            }

            this._writer.WriteEndElement(); //anim
        }

        public void writeMotion(ExtTimeNodeContainer c, int targetRun)
        {
            var tna = c.FirstChildWithType<TimeNodeAtom>();
            var bc = c.FirstChildWithType<TimeMotionBehaviorContainer>();

            var tbc = bc.FirstChildWithType<TimeBehaviorContainer>();
            var tba = tbc.FirstChildWithType<TimeBehaviorAtom>();
            var tmba = bc.FirstChildWithType<TimeMotionBehaviorAtom>();

            var attrNames = new List<string>();
            foreach (var attrName in tbc.FirstChildWithType<TimeStringListContainer>().AllChildrenWithType<TimeVariantValue>())
            {
                attrNames.Add(attrName.stringValue);   
            }

            this._writer.WriteStartElement("p", "animMotion", OpenXmlNamespaces.PresentationML);

            if (tmba.fOriginPropertyUsed)
            {
                switch (tmba.behaviorOrigin)
                {
                    case 0: case 1:
                        this._writer.WriteAttributeString("origin", "parent");
                        break;
                    case 2:
                        this._writer.WriteAttributeString("origin", "layout");
                        break;
                }
            }
            else
            {
                //default
                this._writer.WriteAttributeString("origin", "layout");
            }

            if (tmba.fPathPropertyUsed)
            {
                this._writer.WriteAttributeString("path", bc.FirstChildWithType<TimeVariantValue>().stringValue);
            }

            this._writer.WriteAttributeString("pathEditMode", "relative"); //TODO

            if (tmba.fPointsTypesPropertyUsed)
            {
                if (tbc.FirstChildWithType<TimePropertyList4TimeNodeContainer>() != null)
                {
                    foreach (var v in tbc.FirstChildWithType<TimePropertyList4TimeNodeContainer>().AllChildrenWithType<TimeVariantValue>())
                    {
                        if (v.type == TimeVariantTypeEnum.String)
                        {
                            this._writer.WriteAttributeString("ptsTypes", v.stringValue);
                        }
                        break;
                    }        
                }
                else
                {
                    this._writer.WriteAttributeString("ptsTypes", "");
                }
            }

            this._writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);

            if (tba.fAdditivePropertyUsed)
            {
                switch (tba.behaviorAdditive)
                {
                    case 0: //override
                        this._writer.WriteAttributeString("additive", "base");
                        break;
                    case 1: //add
                        this._writer.WriteAttributeString("additive", "sum");
                        break;
                }
            }

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());

            if (tna.fDurationProperty)
            {
                this._writer.WriteAttributeString("dur", tna.duration.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("dur", "2000");
            }

            if (tna.fFillProperty)
            {
                switch (tna.fill)
                {
                    case 0:
                    case 3:
                        this._writer.WriteAttributeString("fill", "hold");
                        break;
                    case 1:
                    case 4:
                        this._writer.WriteAttributeString("fill", "reset");
                        break;
                    case 2:
                        this._writer.WriteAttributeString("fill", "freeze"); //TODO:verify
                        break;
                }
            }


            this._writer.WriteEndElement(); //cTn

            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("spid", getShapeId(c.FirstDescendantWithType<VisualShapeAtom>().shapeIdRef));
            CheckAndWriteStartEndRuns(c, ref targetRun);

            this._writer.WriteEndElement(); //spTgt
            this._writer.WriteEndElement(); //tgtEl

            this._writer.WriteStartElement("p", "attrNameLst", OpenXmlNamespaces.PresentationML);

            foreach (string attrName in attrNames)
            {
                this._writer.WriteElementString("p", "attrName", OpenXmlNamespaces.PresentationML, attrName);
            }

            this._writer.WriteEndElement(); //attrNameLst

            this._writer.WriteEndElement(); //cBhvr

            this._writer.WriteEndElement(); //anim
        }

        public void writeScale(ExtTimeNodeContainer c, int targetRun)
        {
            var tna = c.FirstChildWithType<TimeNodeAtom>();
            var bc = c.FirstChildWithType<TimeScaleBehaviorContainer>();

            var tbc = bc.FirstChildWithType<TimeBehaviorContainer>();
            var tba = tbc.FirstChildWithType<TimeBehaviorAtom>();
            var tsba = bc.FirstChildWithType<TimeScaleBehaviorAtom>();

            this._writer.WriteStartElement("p", "animScale", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);

            if (tba.fAdditivePropertyUsed)
            {
                switch (tba.behaviorAdditive)
                {
                    case 0: //override
                        this._writer.WriteAttributeString("additive", "base");
                        break;
                    case 1: //add
                        this._writer.WriteAttributeString("additive", "sum");
                        break;
                }
            }

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());

            if (tna.fDurationProperty)
            {
                this._writer.WriteAttributeString("dur", tna.duration.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("dur", "2000");
            }

            if (tna.fFillProperty)
            {
                switch (tna.fill)
                {
                    case 0:
                    case 3:
                        this._writer.WriteAttributeString("fill", "hold");
                        break;
                    case 1:
                    case 4:
                        this._writer.WriteAttributeString("fill", "reset");
                        break;
                    case 2:
                        this._writer.WriteAttributeString("fill", "freeze"); //TODO:verify
                        break;
                }
            }

            this._writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("delay", "650");
            this._writer.WriteEndElement(); //stCondLst
            this._writer.WriteEndElement(); //stCondLst

            this._writer.WriteEndElement(); //cTn

            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("spid", getShapeId(c.FirstDescendantWithType<VisualShapeAtom>().shapeIdRef));
            CheckAndWriteStartEndRuns(c, ref targetRun);

            this._writer.WriteEndElement(); //spTgt
            this._writer.WriteEndElement(); //tgtEl

            this._writer.WriteEndElement(); //cBhvr

            if (tsba.fToPropertyUsed)
            {
                this._writer.WriteStartElement("p", "to", OpenXmlNamespaces.PresentationML);

                this._writer.WriteAttributeString("x", (tsba.fXTo * 1000).ToString());

                this._writer.WriteAttributeString("y", (tsba.fYTo * 1000).ToString());

                this._writer.WriteEndElement(); //to
            }

            this._writer.WriteEndElement(); //animScale
        }

        public void writeColor(RegularContainer c, int targetRun)
        {
            var tna = c.FirstChildWithType<TimeNodeAtom>();
            var bc = c.FirstChildWithType<TimeColorBehaviorContainer>();

            var tbc = bc.FirstChildWithType<TimeBehaviorContainer>();
            var tba = tbc.FirstChildWithType<TimeBehaviorAtom>();
            var tcba = bc.FirstChildWithType<TimeColorBehaviorAtom>();

            var attrNames = new List<string>();
            foreach (var attrName in tbc.FirstChildWithType<TimeStringListContainer>().AllChildrenWithType<TimeVariantValue>())
            {
                attrNames.Add(attrName.stringValue);
            }

            this._writer.WriteStartElement("p", "animClr", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("clrSpc", "rgb");
            this._writer.WriteAttributeString("dir", "cw");

            this._writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("override", "childStyle");

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);

            if (tna.fDurationProperty)
            {
                this._writer.WriteAttributeString("dur", tna.duration.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("dur", "2000");
            }

            if (tna.fFillProperty)
            {
                switch (tna.fill)
                {
                    case 0:
                    case 3:
                        this._writer.WriteAttributeString("fill", "hold");
                        break;
                    case 1:
                    case 4:
                        this._writer.WriteAttributeString("fill", "reset");
                        break;
                    case 2:
                        this._writer.WriteAttributeString("fill", "freeze"); //TODO:verify
                        break;
                }
            }
            else
            {
                this._writer.WriteAttributeString("fill", "hold");
            }

            this._writer.WriteAttributeString("display", "0");
            this._writer.WriteAttributeString("masterRel", "nextClick");
            this._writer.WriteAttributeString("afterEffect", "1");
            this._writer.WriteEndElement(); //cTn

            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("spid", getShapeId(c.FirstDescendantWithType<VisualShapeAtom>().shapeIdRef));

            CheckAndWriteStartEndRuns(c, ref targetRun);

            this._writer.WriteEndElement(); //spTgt

            this._writer.WriteEndElement(); //tgtEl

            this._writer.WriteStartElement("p", "attrNameLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteElementString("p", "attrName", OpenXmlNamespaces.PresentationML, attrNames[0]);

            this._writer.WriteEndElement(); //attrNameLst

            this._writer.WriteEndElement(); //cBhvr


            this._writer.WriteStartElement("p", "to", OpenXmlNamespaces.PresentationML);

            switch (tcba.colorTo.model)
            {
                case 0: //RGB
                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    string val = tcba.colorTo.val1.ToString("X").PadLeft(2, '0') + tcba.colorTo.val2.ToString("X").PadLeft(2, '0') + tcba.colorTo.val3.ToString("X").PadLeft(2, '0');
                    this._writer.WriteAttributeString("val", val);
                    this._writer.WriteEndElement(); //srgbClr
                    break;
                case 1: //HSL
                    break;
                case 2: //scheme
                    this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                    switch (tcba.colorTo.val1)
                    {
                        case 0x00:
                            this._writer.WriteAttributeString("val", "bg1"); //background
                            break;
                        case 0x01:
                            this._writer.WriteAttributeString("val", "tx1"); //text
                            break;
                        case 0x02:
                            this._writer.WriteAttributeString("val", "dk1"); //shadow
                            break;
                        case 0x03:
                            this._writer.WriteAttributeString("val", "tx1"); //title text
                            break;
                        case 0x04:
                            this._writer.WriteAttributeString("val", "bg2"); //fill
                            break;
                        case 0x05:
                            this._writer.WriteAttributeString("val", "accent1"); //accent1
                            break;
                        case 0x06:
                            this._writer.WriteAttributeString("val", "accent2"); //accent2
                            break;
                        case 0x07:
                            this._writer.WriteAttributeString("val", "accent3"); //accent3
                            break;
                    }

                    this._writer.WriteEndElement(); //srgbClr
                    break;

            }



            this._writer.WriteEndElement(); //to

            this._writer.WriteEndElement(); //animClr
        }
        

        public void writeAnimations(ExtTimeNodeContainer c, int targetRun)
        {
            foreach (var c2 in c.AllChildrenWithType<ExtTimeNodeContainer>())
            {
                if (c2.FirstChildWithType<TimeAnimateBehaviourContainer>() != null)
                {
                    writeAnim(c2, targetRun);
                }
                else if (c2.FirstChildWithType<TimeSetBehaviourContainer>() != null)
                {
                    writeSet(c2, ref targetRun);
                }
                else if (c2.FirstChildWithType<TimeEffectBehaviorContainer>() != null)
                {
                    writeAnimEffect(c2, targetRun);
                }
                else if (c2.FirstChildWithType<TimeMotionBehaviorContainer>() != null)
                {
                    writeMotion(c2, targetRun);
                }
                else if (c2.FirstChildWithType<TimeColorBehaviorContainer>() != null)
                {
                    writeColor(c2, targetRun);
                }
                else if (c2.FirstChildWithType<TimeScaleBehaviorContainer>() != null)
                {
                    writeScale(c2, targetRun);
                }
            }
        }

        public void writeAnimEffect(ExtTimeNodeContainer c, int targetRun)
        {
            var tna = c.FirstChildWithType<TimeNodeAtom>();
            var tebc = c.FirstChildWithType<TimeEffectBehaviorContainer>();
            var teba = tebc.FirstChildWithType<TimeEffectBehaviorAtom>();
            var Attributes = new Dictionary<TimePropertyID4TimeNode, TimeVariantValue>();
            
            foreach (var tv in ((RegularContainer)c.ParentRecord).FirstChildWithType<TimePropertyList4TimeNodeContainer>().AllChildrenWithType<TimeVariantValue>())
            {
                Attributes.Add((TimePropertyID4TimeNode)tv.Instance, tv);
            }

            string filter = tebc.FirstChildWithType<TimeVariantValue>().stringValue;

            this._writer.WriteStartElement("p", "animEffect", OpenXmlNamespaces.PresentationML);

            if (Attributes[TimePropertyID4TimeNode.EffectType].intValue == 2)
            {
                this._writer.WriteAttributeString("transition", "out");
            }
            else
            {
                this._writer.WriteAttributeString("transition", "in");
            }

            this._writer.WriteAttributeString("filter", filter);
            this._writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);

            //_writer.WriteAttributeString("additive", "repl");

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());

            if (tna.fDurationProperty)
            {
                this._writer.WriteAttributeString("dur", tna.duration.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("dur", "500");
            }

            this._writer.WriteStartElement("p", "stCondLst", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cond", OpenXmlNamespaces.PresentationML);

            if (c.FirstChildWithType<TimeConditionContainer>() != null)
            {
                this._writer.WriteAttributeString("delay", c.FirstChildWithType<TimeConditionContainer>().FirstChildWithType<TimeConditionAtom>().delay.ToString());
            }
            else
            {
                this._writer.WriteAttributeString("delay", "0");
            }

            this._writer.WriteEndElement(); //cond

            this._writer.WriteEndElement(); //stCondLst

            this._writer.WriteEndElement(); //cTn
            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("spid", getShapeId(c.FirstDescendantWithType<VisualShapeAtom>().shapeIdRef));
            
            CheckAndWriteStartEndRuns((ExtTimeNodeContainer)c.ParentRecord, ref targetRun);

            this._writer.WriteEndElement(); //spTgt
            this._writer.WriteEndElement(); //tgtEl
            this._writer.WriteEndElement(); //cBhvr
            this._writer.WriteEndElement(); //animEffect
        }

        private void CheckAndWriteStartEndRuns(RegularContainer c, ref int targetRun)
        {
            var vsa = c.FirstDescendantWithType<VisualShapeAtom>(); 

            if (!this._stm.spidToId.ContainsKey((int)vsa.shapeIdRef))
            {
                return;
            }

            
            if (vsa.type == TimeVisualElementEnum.TextRange)
            {
                int i = 0;
                foreach (var p in this.TextAreasForAnimation)
                {
                    if (p.X <= vsa.data1 && p.Y >= vsa.data2)
                    {
                        targetRun = i;
                        break;
                    }
                    i++;
                }
                if (targetRun == -1)
                {
                    if (vsa.data1 > 0 && this.TextAreasForAnimation.Count == 0) this.TextAreasForAnimation.Add(new Point(0, vsa.data1));
                    this.TextAreasForAnimation.Add(new Point(vsa.data1, vsa.data2));
                    targetRun = this.TextAreasForAnimation.Count - 1;
                }

                this._writer.WriteStartElement("p", "txEl", OpenXmlNamespaces.PresentationML);
                this._writer.WriteStartElement("p", "pRg", OpenXmlNamespaces.PresentationML);
                this._writer.WriteAttributeString("st", targetRun.ToString());
                this._writer.WriteAttributeString("end", targetRun.ToString());
                this._writer.WriteEndElement(); //pRg
                this._writer.WriteEndElement(); //txEl
            }
        }

        public void writeAnimEffect(AnimationInfoAtom animinfo, ExtTimeNodeContainer c, int targetRun)
        {
            this._writer.WriteStartElement("p", "animEffect", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("transition", "in");

            switch (animinfo.animEffect)
            {
                case 0x00: //Cut
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //not through black
                        case 0x02: //same as 0x00
                            this._writer.WriteAttributeString("filter", "cut(false)");
                            break;
                        case 0x01: //through black
                            this._writer.WriteAttributeString("filter", "cut(true)");
                            break;
                    }
                    break;
                case 0x01: //Random
                    this._writer.WriteAttributeString("filter", "random");
                    break;
                case 0x02: //Blinds
                    if (animinfo.animEffectDirection == 0x01)
                    {
                        this._writer.WriteAttributeString("filter", "blinds(horizontal)");
                    }
                    else
                    {
                        this._writer.WriteAttributeString("filter", "blinds(vertical)");
                    }
                    break;
                case 0x03: //Checker
                    if (animinfo.animEffectDirection == 0x00)
                    {
                        this._writer.WriteAttributeString("filter", "checkerboard(across)");
                    }
                    else
                    {
                        this._writer.WriteAttributeString("filter", "checkerboard(down)");
                    }
                    break;
                case 0x04: //Cover
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //r->l
                            this._writer.WriteAttributeString("filter", "cover(l)");
                            break;
                        case 0x01: //b->t
                            this._writer.WriteAttributeString("filter", "cover(u)");
                            break;
                        case 0x02: //l->r
                            this._writer.WriteAttributeString("filter", "cover(r)");
                            break;
                        case 0x03: //t->b
                            this._writer.WriteAttributeString("filter", "cover(d)");
                            break;
                        case 0x04: //br->tl
                            this._writer.WriteAttributeString("filter", "cover(lu)");
                            break;
                        case 0x05: //bl->tr
                            this._writer.WriteAttributeString("filter", "cover(ru)");
                            break;
                        case 0x06: //tr->bl
                            this._writer.WriteAttributeString("filter", "cover(ld)");
                            break;
                        case 0x07: //tl->br
                            this._writer.WriteAttributeString("filter", "cover(rd)");
                            break;
                    }
                    break;
                case 0x05: //Dissolve
                    this._writer.WriteAttributeString("filter", "dissolve");
                    break;
                case 0x06: //Fade
                    this._writer.WriteAttributeString("filter", "fade");
                    break;
                case 0x07: //Pull
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //r->l
                            this._writer.WriteAttributeString("filter", "pull(l)");
                            break;
                        case 0x01: //b->t
                            this._writer.WriteAttributeString("filter", "pull(u)");
                            break;
                        case 0x02: //l->r
                            this._writer.WriteAttributeString("filter", "pull(r)");
                            break;
                        case 0x03: //t->b
                            this._writer.WriteAttributeString("filter", "pull(d)");
                            break;
                        case 0x04: //br->tl
                            this._writer.WriteAttributeString("filter", "pull(lu)");
                            break;
                        case 0x05: //bl->tr
                            this._writer.WriteAttributeString("filter", "pull(ru)");
                            break;
                        case 0x06: //tr->bl
                            this._writer.WriteAttributeString("filter", "pull(ld)");
                            break;
                        case 0x07: //tl->br
                            this._writer.WriteAttributeString("filter", "pull(rd)");
                            break;
                    }
                    break;
                case 0x08: //Random bar
                    if (animinfo.animEffectDirection == 0x01)
                    {
                        this._writer.WriteAttributeString("filter", "randomBar(horz)");
                    }
                    else
                    {
                        this._writer.WriteAttributeString("filter", "randomBar(vert)");
                    }
                    break;
                case 0x09: //Strips
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x04: //br->tl
                            this._writer.WriteAttributeString("filter", "strips(lu)");
                            break;
                        case 0x05: //bl->tr
                            this._writer.WriteAttributeString("filter", "strips(ru)");
                            break;
                        case 0x06: //tr->bl
                            this._writer.WriteAttributeString("filter", "strips(ld)");
                            break;
                        case 0x07: //tl->br
                            this._writer.WriteAttributeString("filter", "strips(rd)");
                            break;
                    }
                    break;
                case 0x0a: //Wipe
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //r->l
                            this._writer.WriteAttributeString("filter", "wipe(l)");
                            break;
                        case 0x01: //b->t
                            this._writer.WriteAttributeString("filter", "wipe(u)");
                            break;
                        case 0x02: //l->r
                            this._writer.WriteAttributeString("filter", "wipe(r)");
                            break;
                        case 0x03: //t->b
                            this._writer.WriteAttributeString("filter", "wipe(d)");
                            break;
                    }
                    break;
                case 0x0b: //Zoom (box)
                    if (animinfo.animEffectDirection == 0x00)
                    {
                        this._writer.WriteAttributeString("filter", "box(out)");
                    }
                    else
                    {
                        this._writer.WriteAttributeString("filter", "box(in)");
                    }
                    break;
                case 0x0c: //Fly
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //from left
                            this._writer.WriteAttributeString("filter", "slide(fromLeft)");
                            break;
                        case 0x01: //from top
                            this._writer.WriteAttributeString("filter", "slide(fromTop)");
                            break;
                        case 0x02: //from right
                            this._writer.WriteAttributeString("filter", "slide(fromRight)");
                            break;
                        case 0x03: //from bottom  
                            this._writer.WriteAttributeString("filter", "slide(fromBottom)");
                            break;
                        case 0x04: //from top left
                        case 0x05: //from top right
                        case 0x06: //from bottom left
                        case 0x07: //from bottom right
                        case 0x08: //from left edge of shape / text
                        case 0x09: //from bottom edge of shape / text
                        case 0x0a: //from right edge of shape / text
                        case 0x0b: //from top edge of shape / text
                        case 0x0c: //crawl from left
                        case 0x0d: //crawl from top 
                        case 0x0e: //crawl from right
                        case 0x0f: //crawl from bottom
                        case 0x10: //zoom 0 -> 1
                        case 0x11: //zoom 0.5 -> 1
                        case 0x12: //zoom 4 -> 1
                        case 0x13: //zoom 1.5 -> 1
                        case 0x14: //zoom 0 -> 1; screen center -> actual center
                        case 0x15: //zoom 4 -> 1; bottom -> actual position
                        case 0x16: //stretch center -> l & r
                        case 0x17: //stretch l -> r
                        case 0x18: //stretch t -> b
                        case 0x19: //stretch r -> l
                        case 0x1a: //stretch b -> t
                        case 0x1b: //rotate around vertical axis that passes through its center
                        case 0x1c: //flies in a spiral
                            this._writer.WriteAttributeString("filter", "slide(fromBottom)");
                            break;
                    }
                    break;
                case 0x0d: //Split
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //horz m -> tb
                            this._writer.WriteAttributeString("filter", "split(outHorizontal)");
                            break;
                        case 0x01: //horz tb -> m
                            this._writer.WriteAttributeString("filter", "split(inHorizontal)");
                            break;
                        case 0x02: //vert m -> lr
                            this._writer.WriteAttributeString("filter", "split(outVertical)");
                            break;
                        case 0x03: //vert
                            this._writer.WriteAttributeString("filter", "split(inVertical)");
                            break;
                    }
                    break;
                case 0x0e: //Flash
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x00: //after short time
                        case 0x01: //after medium time
                        case 0x02: //after long time
                            break;
                    }
                    break;
                case 0x0f:
                case 0x11: //Diamond
                    this._writer.WriteAttributeString("filter", "diamond(out)");
                    break;
                case 0x12: //Plus
                    this._writer.WriteAttributeString("filter", "plus");
                    break;
                case 0x13: //Wedge
                    this._writer.WriteAttributeString("filter", "wedge");
                    break;
                case 0x14:
                case 0x15:
                case 0x16:
                case 0x17:
                case 0x18:
                case 0x19:
                case 0x1a: //Wheel
                    switch (animinfo.animEffectDirection)
                    {
                        case 0x01: //1 spoke
                            this._writer.WriteAttributeString("filter", "wheel(1)");
                            break;
                        case 0x02: //2 spokes
                            this._writer.WriteAttributeString("filter", "wheel(2)");
                            break;
                        case 0x03: //3 spokes
                            this._writer.WriteAttributeString("filter", "wheel(3)");
                            break;
                        case 0x04: //4 spokes
                            this._writer.WriteAttributeString("filter", "wheel(4)");
                            break;
                        case 0x08: //8 spokes
                            this._writer.WriteAttributeString("filter", "wheel(8)");
                            break;
                    }
                    break;
                case 0x1b: //Circle
                    this._writer.WriteAttributeString("filter", "circle");
                    break;
                default:
                    this._writer.WriteAttributeString("filter", "blinds(horizontal)");
                    break;
            }

            this._writer.WriteStartElement("p", "cBhvr", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "cTn", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("id", (++this.lastID).ToString());
            this._writer.WriteAttributeString("dur", "500");
            this._writer.WriteEndElement(); //cTn

            this._writer.WriteStartElement("p", "tgtEl", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("p", "spTgt", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("spid", getShapeId(c.FirstDescendantWithType<VisualShapeAtom>().shapeIdRef));

            CheckAndWriteStartEndRuns(c, ref targetRun);

            this._writer.WriteEndElement(); //spTgt

            this._writer.WriteEndElement(); //tgtEl

            this._writer.WriteEndElement(); //cBhvr

            this._writer.WriteEndElement(); //animEffect
        }
    }
}
