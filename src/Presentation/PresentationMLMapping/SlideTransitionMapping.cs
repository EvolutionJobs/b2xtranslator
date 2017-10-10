

using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using System.Drawing;

namespace b2xtranslator.PresentationMLMapping
{
    class SlideTransitionMapping :
        AbstractOpenXmlMapping//,
    {
        protected ConversionContext _ctx;
        private ShapeTreeMapping _stm;
        private List<Point> TextAreasForAnimation = new List<Point>();

        public SlideTransitionMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            this._ctx = ctx;
        }

        public void Apply(SlideShowSlideInfoAtom slideshow)
        {
            if (slideshow.fAutoAdvance)
            {
                this._writer.WriteStartElement("p", "transition", OpenXmlNamespaces.PresentationML);
                this._writer.WriteAttributeString("advTm", slideshow.slideTime.ToString());

                switch (slideshow.speed)
                {
                    case 0:
                        this._writer.WriteAttributeString("spd", "slow");
                        break;
                    case 1:
                        this._writer.WriteAttributeString("spd", "med");
                        break;
                    case 2:
                        this._writer.WriteAttributeString("spd", "fast");
                        break;
                }

                switch (slideshow.effectType)
                {
                    case 0:
                        //simple cut
                        if (slideshow.effectDirection == 1)
                        {
                            this._writer.WriteStartElement("p", "cut", OpenXmlNamespaces.PresentationML);
                            this._writer.WriteAttributeString("thruBlk", "true");
                            this._writer.WriteEndElement();
                        }
                        break;
                    case 1: //random
                        this._writer.WriteElementString("p", "random", OpenXmlNamespaces.PresentationML, "");
                        break;
                    case 2: //blinds
                        this._writer.WriteStartElement("p", "blinds", OpenXmlNamespaces.PresentationML);
                        switch (slideshow.effectDirection)
                        {
                            case 0:
                                this._writer.WriteAttributeString("dir", "vert");
                                break;
                            case 1:
                                this._writer.WriteAttributeString("dir", "horz");
                                break;
                        }
                        this._writer.WriteEndElement();
                        break;
                    case 3:
                        this._writer.WriteStartElement("p","checker", OpenXmlNamespaces.PresentationML);
                        switch (slideshow.effectDirection)
                        {
                            case 0:
                                this._writer.WriteAttributeString("dir", "horz");
                                break;
                            case 1:
                                this._writer.WriteAttributeString("dir","vert");
                                break;
                        }
                        this._writer.WriteEndElement();
                        break;
                    case 4: //cover
                        this._writer.WriteStartElement("p", "cover", OpenXmlNamespaces.PresentationML);
                        switch (slideshow.effectDirection)
                        {
                            case 0:
                                this._writer.WriteAttributeString("dir", "l");
                                break;
                            case 1:
                                this._writer.WriteAttributeString("dir", "u");
                                break;
                            case 2:
                                this._writer.WriteAttributeString("dir", "r");
                                break;
                            case 3:
                                this._writer.WriteAttributeString("dir", "d");
                                break;
                            case 4:
                                this._writer.WriteAttributeString("dir", "lu");
                                break;
                            case 5:
                                this._writer.WriteAttributeString("dir", "ru");
                                break;
                            case 6:
                                this._writer.WriteAttributeString("dir", "ld");
                                break;
                            case 7:
                                this._writer.WriteAttributeString("dir", "rd");
                                break;
                        }
                        this._writer.WriteEndElement();
                        break;
                    case 5: //dissolve
                        this._writer.WriteElementString("p", "dissolve", OpenXmlNamespaces.PresentationML, "");
                        break;
                    case 6: //fade
                        this._writer.WriteElementString("p", "fade", OpenXmlNamespaces.PresentationML, "");
                         break;
                    case 7: //uncover
                        this._writer.WriteStartElement("p", "push", OpenXmlNamespaces.PresentationML); //TODO
                         switch (slideshow.effectDirection)
                         {
                             case 0:
                                this._writer.WriteAttributeString("dir", "l");
                                 break;
                             case 1:
                                this._writer.WriteAttributeString("dir", "u");
                                 break;
                             case 2:
                                this._writer.WriteAttributeString("dir", "r");
                                 break;
                             case 3:
                                this._writer.WriteAttributeString("dir", "d");
                                 break;
                             case 4:
                                this._writer.WriteAttributeString("dir", "lu");
                                 break;
                             case 5:
                                this._writer.WriteAttributeString("dir", "ru");
                                 break;
                             case 6:
                                this._writer.WriteAttributeString("dir", "ld");
                                 break;
                             case 7:
                                this._writer.WriteAttributeString("dir", "rd");
                                 break;
                         }
                        this._writer.WriteEndElement();
                        break;
                    case 8: //random bars
                        this._writer.WriteStartElement("p", "randomBar", OpenXmlNamespaces.PresentationML); 
                        switch (slideshow.effectDirection)
                        {
                            case 0:
                                this._writer.WriteAttributeString("dir", "horz");
                                break;
                            case 1:
                                this._writer.WriteAttributeString("dir", "vert");
                                break;                            
                        }
                        this._writer.WriteEndElement();
                        break;
                    case 9: //strips
                        this._writer.WriteStartElement("p", "strips", OpenXmlNamespaces.PresentationML);
                        switch (slideshow.effectDirection)
                        {
                            case 4:
                                this._writer.WriteAttributeString("dir", "lu");
                                break;
                            case 5:
                                this._writer.WriteAttributeString("dir", "ru");
                                break;
                            case 6:
                                this._writer.WriteAttributeString("dir", "ld");
                                break;
                            case 7:
                                this._writer.WriteAttributeString("dir", "rd");
                                break;
                        }
                        this._writer.WriteEndElement();
                        break;
                    case 10: //wipe
                        this._writer.WriteStartElement("p", "wipe", OpenXmlNamespaces.PresentationML); 
                        switch (slideshow.effectDirection)
                        {
                            case 0:
                                this._writer.WriteAttributeString("dir", "l");
                                break;
                            case 1:
                                this._writer.WriteAttributeString("dir", "u");
                                break;
                            case 2:
                                this._writer.WriteAttributeString("dir", "r");
                                break;
                            case 3:
                                this._writer.WriteAttributeString("dir", "d");
                                break;                         
                        }
                        this._writer.WriteEndElement();
                        break;
                    case 11: // box in/out
                        this._writer.WriteStartElement("p", "zoom", OpenXmlNamespaces.PresentationML); //TODO
                        switch (slideshow.effectDirection)
                        {
                            case 0:
                                this._writer.WriteAttributeString("dir", "out");
                                break;
                            case 1:
                                this._writer.WriteAttributeString("dir", "in");
                                break;
                        }
                        this._writer.WriteEndElement();
                        break;
                    case 13: //split
                        this._writer.WriteStartElement("p", "split", OpenXmlNamespaces.PresentationML); 
                        switch (slideshow.effectDirection)
                        {
                            case 0:
                                this._writer.WriteAttributeString("dir", "out");
                                this._writer.WriteAttributeString("orient", "horz");
                                break;
                            case 1:
                                this._writer.WriteAttributeString("dir", "in");
                                this._writer.WriteAttributeString("orient", "horz");
                                break;
                            case 2:
                                this._writer.WriteAttributeString("dir", "out");
                                this._writer.WriteAttributeString("orient", "vert");
                                break;
                            case 3:
                                this._writer.WriteAttributeString("dir", "in");
                                this._writer.WriteAttributeString("orient", "vert");
                                break;
                        }
                        this._writer.WriteEndElement();
                        break;
                    case 17:
                        this._writer.WriteElementString("p", "diamond", OpenXmlNamespaces.PresentationML,"");
                        break;
                    case 18: //plus
                        this._writer.WriteElementString("p", "plus", OpenXmlNamespaces.PresentationML, "");
                        break;
                    case 19: //wedge
                        this._writer.WriteElementString("p", "wedge", OpenXmlNamespaces.PresentationML, "");
                        break;
                    case 20: //push
                        this._writer.WriteStartElement("p", "push", OpenXmlNamespaces.PresentationML);
                        switch (slideshow.effectDirection)
                        {
                            case 0:
                                this._writer.WriteAttributeString("dir", "l");
                                break;
                            case 1:
                                this._writer.WriteAttributeString("dir", "u");
                                break;
                            case 2:
                                this._writer.WriteAttributeString("dir", "r");
                                break;
                            case 3:
                                this._writer.WriteAttributeString("dir", "d");
                                break;
                        }
                        this._writer.WriteEndElement();
                        break;
                    case 21: //comb
                        this._writer.WriteStartElement("p", "comb", OpenXmlNamespaces.PresentationML);
                        switch (slideshow.effectDirection)
                        {
                            case 0:
                                this._writer.WriteAttributeString("dir", "horz");
                                break;
                            case 1:
                                this._writer.WriteAttributeString("dir", "vert");
                                break;
                        }
                        this._writer.WriteEndElement();
                        break;
                    case 22: //newsflash
                        this._writer.WriteElementString("p", "newsflash", OpenXmlNamespaces.PresentationML, "");
                        break;
                    case 23: //alphafade TODO
                        break;
                    case 26: //wheel
                        this._writer.WriteStartElement("p", "wheel", OpenXmlNamespaces.PresentationML); //TODO
                        this._writer.WriteAttributeString("spokes", slideshow.effectDirection.ToString());
                        this._writer.WriteEndElement();
                        break;
                    case 27: //circle
                        this._writer.WriteElementString("p", "circle", OpenXmlNamespaces.PresentationML, "");
                        break;
                    default:
                        break;
                }

                this._writer.WriteEndElement();
            }
        }

        
    }
}
