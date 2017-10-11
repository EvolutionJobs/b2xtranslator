

using System;
using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using System.Xml;
using System.Reflection;
using b2xtranslator.Tools;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PresentationMLMapping
{
    public static class Utils
    {
        private static readonly double MC_PER_EMU = 1587.5;

        public static int MasterCoordToEMU(int mc)
        {
            return (int) (mc * MC_PER_EMU);
        }

        public static int EMUToMasterCoord(int emu)
        {
            return (int) (emu / MC_PER_EMU);
        }
                
        public static XmlDocument GetDefaultDocument(string filename)
        {
            var a = Assembly.GetExecutingAssembly();
            var s = a.GetManifestResourceStream(string.Format("{0}.Defaults.{1}.xml",
                typeof(Utils).Namespace, filename));

            var doc = new XmlDocument();
            doc.Load(s);
            return doc;
        }

        public static string SlideSizeTypeToXMLValue(SlideSizeType sst)
        {
            // OOXML Spec § 4.8.22
            switch (sst)
            {
                case SlideSizeType.A4Paper:
                    return "A4";

                case SlideSizeType.Banner:
                    return "banner";

                case SlideSizeType.Custom:
                    return "custom";

                case SlideSizeType.LetterSizedPaper:
                    return "letter";

                case SlideSizeType.OnScreen:
                    return "screen4x3";

                case SlideSizeType.Overhead:
                    return "overhead";

                case SlideSizeType.Size35mm:
                    return "35mm";

                default:
                    throw new NotImplementedException(
                        string.Format("Can't convert slide size type {0} to XML value", sst));
            }
        }

        public static string PlaceholderIdToXMLValue(PlaceholderEnum pid)
        {
            switch (pid)
            {
                case PlaceholderEnum.MasterDate:
                    return "dt";

                case PlaceholderEnum.MasterSlideNumber:
                    return "sldNum";

                case PlaceholderEnum.MasterFooter:
                    return "ftr";

                case PlaceholderEnum.MasterHeader:
                    return "hdr";

                case PlaceholderEnum.MasterTitle:
                case PlaceholderEnum.Title:
                    return "title";

                case PlaceholderEnum.MasterBody:
                case PlaceholderEnum.Body:
                case PlaceholderEnum.NotesBody:
                case PlaceholderEnum.MasterNotesBody:
                    return "body";

                case PlaceholderEnum.MasterCenteredTitle:
                case PlaceholderEnum.CenteredTitle:
                    return "ctrTitle";

                case PlaceholderEnum.MasterSubtitle:
                case PlaceholderEnum.Subtitle:
                    return "subTitle";

                case PlaceholderEnum.ClipArt:
                    return "clipArt";

                case PlaceholderEnum.Graph:
                    return "chart";

                case PlaceholderEnum.OrganizationChart:
                    return "dgm";

                case PlaceholderEnum.MediaClip:
                    return "media";

                case PlaceholderEnum.Table:
                    return "tbl";

                case PlaceholderEnum.NotesSlideImage:
                case PlaceholderEnum.MasterNotesSlideImage:
                    return "sldImg";


                default:
                    throw new NotImplementedException("Don't know how to map placeholder id " + pid);
            }
        }

        public static string SlideLayoutTypeToFilename(SlideLayoutType type, PlaceholderEnum[] placeholderTypes)
        {
            switch (type)
            {
                case SlideLayoutType.BigObject:
                    return "objOnly";

                case SlideLayoutType.Blank:
                    return "blank";

                case SlideLayoutType.FourObjects:
                    return "fourObj";

                case SlideLayoutType.TitleAndBody:
                    {
                        var body = placeholderTypes[1];

                        if (body == PlaceholderEnum.Table)
                        {
                            return "tbl";
                        }
                        else if (body == PlaceholderEnum.OrganizationChart)
                        {
                            return "dgm";
                        }
                        else if (body == PlaceholderEnum.Graph)
                        {
                            return "chart";
                        }
                        else
                        {
                            return "obj";
                        }
                    }

                case SlideLayoutType.TitleOnly:
                    return "titleOnly";

                case SlideLayoutType.TitleSlide:
                    return "title";

                case SlideLayoutType.TwoColumnsAndTitle:
                    {
                        var leftType = placeholderTypes[1];
                        var rightType = placeholderTypes[2];

                        if (leftType == PlaceholderEnum.Body && rightType == PlaceholderEnum.Object)
                        {
                            return "txAndObj";
                        }
                        else if (leftType == PlaceholderEnum.Object && rightType == PlaceholderEnum.Body)
                        {
                            return "objAndTx";
                        }
                        else if (leftType == PlaceholderEnum.Body && rightType == PlaceholderEnum.ClipArt)
                        {
                            return "txAndClipArt";
                        }
                        else if (leftType == PlaceholderEnum.ClipArt && rightType == PlaceholderEnum.Body)
                        {
                            return "clipArtAndTx";
                        }
                        else if (leftType == PlaceholderEnum.Body && rightType == PlaceholderEnum.Graph)
                        {
                            return "txAndChart";
                        }
                        else if (leftType == PlaceholderEnum.Graph && rightType == PlaceholderEnum.Body)
                        {
                            return "chartAndTx";
                        }
                        else if (leftType == PlaceholderEnum.Body && rightType == PlaceholderEnum.MediaClip)
                        {
                            return "txAndMedia";
                        }
                        else if (leftType == PlaceholderEnum.MediaClip && rightType == PlaceholderEnum.Body)
                        {
                            return "mediaAndTx";
                        }
                        else
                        {
                            return "twoObj";
                        }
                    }

                case SlideLayoutType.TwoColumnsLeftTwoRows:
                    {
                        var rightType = placeholderTypes[2];

                        if (rightType == PlaceholderEnum.Object)
                        {
                            return "twoObjAndObj";
                        }
                        else if (rightType == PlaceholderEnum.Body)
                        {
                            return "twoObjAndTx";
                        }
                        else
                        {
                            throw new NotImplementedException(string.Format(
                                "Don't know how to map TwoColumnLeftTwoRows with rightType = {0}",
                                rightType
                            ));
                        }
                    }

                case SlideLayoutType.TwoColumnsRightTwoRows:
                    {
                        var leftType = placeholderTypes[1];

                        if (leftType == PlaceholderEnum.Object)
                        {
                            return "objAndTwoObj";
                        }
                        else if (leftType == PlaceholderEnum.Body)
                        {
                            return "txAndTwoObj";
                        }
                        else
                        {
                            throw new NotImplementedException(string.Format(
                                "Don't know how to map TwoColumnRightTwoRows with leftType = {0}",
                                leftType
                            ));
                        }
                    }

                case SlideLayoutType.TwoRowsAndTitle:
                    {
                        var topType = placeholderTypes[1];
                        var bottomType = placeholderTypes[2];

                        if (topType == PlaceholderEnum.Body && bottomType == PlaceholderEnum.Object)
                        {
                            return "txOverObj";
                        }
                        else if (topType == PlaceholderEnum.Object && bottomType == PlaceholderEnum.Body)
                        {
                            return "objOverTx";
                        }
                        else
                        {
                            throw new NotImplementedException(string.Format(
                                "Don't know how to map TwoRowsAndTitle with topType = {0} and bottomType = {1}",
                                topType, bottomType
                            ));
                        }
                    }

                case SlideLayoutType.TwoRowsTopTwoColumns:
                    return "twoObjOverTx";

                default:
                    throw new NotImplementedException("Don't know how to map slide layout type " + type); 
            }
        }

        public static string getRGBColorFromOfficeArtCOLORREF(uint value, RegularContainer slide, b2xtranslator.OfficeDrawing.ShapeOptions so)
        {
            string dummy = "";
            return getRGBColorFromOfficeArtCOLORREF(value, slide, so, ref dummy);
        }



        public static string getRGBColorFromOfficeArtCOLORREF(uint value, RegularContainer slide, b2xtranslator.OfficeDrawing.ShapeOptions so, ref string SchemeType)
        {
            var bytes = BitConverter.GetBytes(value);
            bool fPaletteIndex = (bytes[3] & 1) != 0;
            bool fPaletteRGB = (bytes[3] & (1 << 1)) != 0;
            bool fSystemRGB = (bytes[3] & (1 << 2)) != 0;
            bool fSchemeIndex = (bytes[3] & (1 << 3)) != 0;
            bool fSysIndex = (bytes[3] & (1 << 4)) != 0;

            var colors = slide.AllChildrenWithType<ColorSchemeAtom>();
            ColorSchemeAtom MasterScheme = null;
            foreach (var color in colors)
            {
                if (color.Instance == 1) MasterScheme = color;
            }

            if (fSysIndex)
            {
                ushort val = BitConverter.ToUInt16(bytes, 0);
                string result = "";
                switch (val & 0x00ff)
                {
                    case 0xF0: //shape fill color
                        if (so.OptionsByID.ContainsKey(b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.fillColor))
                        {
                            result = getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.fillColor].op,slide,so);
                        } else {
                            result = new RGBColor(MasterScheme.Fills, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;  //TODO: find out which color to use in this case
                        }
                        break;
                    case 0xF1: //shape line color if it is a line else shape fill color TODO!!
                        if (so.FirstAncestorWithType<OfficeDrawing.ShapeContainer>().FirstChildWithType<OfficeDrawing.Shape>().Instance == 1)
                        {
                            if (so.OptionsByID.ContainsKey(b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.fillColor))
                            {
                                result = getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.fillColor].op, slide, so);
                            }
                            else
                            {
                                result = new RGBColor(MasterScheme.Fills, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;  //TODO: find out which color to use in this case
                            }
                        }
                        else
                        {
                            if (so.OptionsByID.ContainsKey(b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.lineColor))
                            {
                                result = getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.lineColor].op, slide, so);
                            }
                            else
                            {
                                result = new RGBColor(MasterScheme.TextAndLines, RGBColor.ByteOrder.RedFirst).SixDigitHexCode; //TODO: find out which color to use in this case
                            }
                        }
                        break;
                    case 0xF2: //shape line color
                        if (so.OptionsByID.ContainsKey(b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.lineColor))
                        {
                            result = getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.lineColor].op, slide, so);
                        }
                        else
                        {
                            result = new RGBColor(MasterScheme.TextAndLines, RGBColor.ByteOrder.RedFirst).SixDigitHexCode; //TODO: find out which color to use in this case
                        }
                        break;
                    case 0xF3: //shape shadow color
                        result = getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.shadowColor].op, slide, so);
                        break;
                    case 0xF4: //current or last used color
                    case 0xF5: //shape fill background color
                        result = getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                        break;
                    case 0xF6: //shape line background color
                        result = getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.lineBackColor].op, slide, so);
                        break;
                    case 0xF7: //shape fill color if shape contains a fill else line color
                        result = getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.fillColor].op, slide, so);
                        break;
                    case 0xFF: //undefined
                        return "";
                }

                if (result.Length == 0) return "";

                int red = int.Parse(result.Substring(0, 2), System.Globalization.NumberStyles.HexNumber);
                int green = int.Parse(result.Substring(2, 2), System.Globalization.NumberStyles.HexNumber);
                int blue = int.Parse(result.Substring(4, 2), System.Globalization.NumberStyles.HexNumber);
                int v = (int)bytes[2];
                //int res;
                return result;
                //switch (val & 0xff00)
                //{
                //    case 0x100:
                //        if (blue == 0xff) return result;
                //        if (blue == 0x00) return "000000";

                //        res = int.Parse(result, System.Globalization.NumberStyles.HexNumber);
                //        if (!so.OptionsByID.ContainsKey(b2xtranslator.OfficeDrawing.ShapeOptions.PropertyId.ShadowStyleBooleanProperties))
                //        res -= v; //this is wrong for shadow17
                //        if (res < 0) res = 0;
                //        return res.ToString("X").PadLeft(6, '0');
                //    case 0x200:
                //        if (blue == 0xff) return result;
                //        if (blue == 0x00) return "FFFFFF";

                //        res = int.Parse(result, System.Globalization.NumberStyles.HexNumber);
                //        res += v;
                //        return res.ToString("X").PadLeft(6, '0');
                //    case 0x300:
                //        red += v;
                //        green += v;
                //        blue += v;
                //        if (red > 0xff) red = 0xff;
                //        if (green > 0xff) green = 0xff;
                //        if (blue > 0xff) blue = 0xff;
                //        return red.ToString("X").PadLeft(2, '0') + green.ToString("X").PadLeft(2, '0') + blue.ToString("X").PadLeft(2, '0');
                //    case 0x400:
                //        red -= v;
                //        green -= v;
                //        blue -= v;
                //        if (red < 0) red = 0x0;
                //        if (green < 0) green = 0x0;
                //        if (blue < 0) blue = 0x0;
                //        return red.ToString("X").PadLeft(2, '0') + green.ToString("X").PadLeft(2, '0') + blue.ToString("X").PadLeft(2, '0');
                //    case 0x500:
                //        red = v - red;
                //        green = v - green;
                //        blue = v - blue;
                //        if (red < 0) red = 0x0;
                //        if (green < 0) green = 0x0;
                //        if (blue < 0) blue = 0x0;
                //        return red.ToString("X").PadLeft(2, '0') + green.ToString("X").PadLeft(2, '0') + blue.ToString("X").PadLeft(2, '0');
                //    default:
                //        break;
                //}
            } 
            
            if (fSchemeIndex)
            {
                //red is the index to the color scheme
                //List<ColorSchemeAtom> colors = slide.AllChildrenWithType<ColorSchemeAtom>();
                //ColorSchemeAtom MasterScheme = null;
                //foreach (ColorSchemeAtom color in colors)
                //{
                //    if (color.Instance == 1) MasterScheme = color;
                //}

                switch (bytes[0])
                {
                    case 0x00: //background
                        SchemeType = "bg1";
                        return new RGBColor(MasterScheme.Background, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                    case 0x01: //text
                        SchemeType = "tx1";
                        return new RGBColor(MasterScheme.TextAndLines, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                    case 0x02: //shadow
                        return new RGBColor(MasterScheme.Shadows, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                    case 0x03: //title
                        return new RGBColor(MasterScheme.TitleText, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                    case 0x04: //fill
                        SchemeType = "accent1";
                        return new RGBColor(MasterScheme.Fills, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                    case 0x05: //accent1
                        SchemeType = "accent2";
                        return new RGBColor(MasterScheme.Accent, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                    case 0x06: //accent2
                        SchemeType = "hlink";
                        return new RGBColor(MasterScheme.AccentAndHyperlink, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                    case 0x07: //accent3
                        return new RGBColor(MasterScheme.AccentAndFollowedHyperlink, RGBColor.ByteOrder.RedFirst).SixDigitHexCode;
                    case 0xFE: //sRGB
                        return bytes[0].ToString("X").PadLeft(2, '0') + bytes[1].ToString("X").PadLeft(2, '0') + bytes[3].ToString("X").PadLeft(2, '0');
                    case 0xFF: //undefined
                        break;
                }                
            }
            return bytes[0].ToString("X").PadLeft(2, '0') + bytes[1].ToString("X").PadLeft(2, '0') + bytes[2].ToString("X").PadLeft(2, '0');
        }

        //public static string getPrstForPattern(string blipNamePattern)
        //{
        //    switch (blipNamePattern)
        //    {
        //        case "5%":
        //            return  "pct5";
        //        case "10%":
        //            return  "pct10";
        //        case "20%":
        //            return  "pct20";
        //        case "25%":
        //            return  "pct25";
        //        case "30%":
        //            return  "pct30";
        //        case "40%":
        //            return  "pct40";
        //        case "50%":
        //            return  "pct50";
        //        case "60%":
        //            return  "pct60";
        //        case "70%":
        //            return  "pct70";
        //        case "75%":
        //            return  "pct75";
        //        case "80%":
        //            return  "pct80";
        //        case "90%":
        //            return  "pct90";
        //        case "dark horizontal":
        //            return  "dkHorz";
        //        case "dark vertical":
        //            return  "dkVert";
        //        case "dark downward diagonal":
        //            return  "dkDnDiag";
        //        case "dark upward diagonal":
        //            return  "dkUpDiag";
        //        case "dashed downward diagonal":
        //            return  "dashDnDiag";
        //        case "dashed horizontal":
        //            return  "dashHorz";
        //        case "dashed vertical":
        //            return  "dashVert";
        //        case "dashed upward diagonal":
        //            return  "dashUpDiag";
        //        case "diagonal brick":
        //            return  "diagBrick";
        //        case "divot":
        //            return  "divot";
        //        case "dotted grid":
        //            return  "dotGrid";
        //        case "dotted diamond":
        //            return  "dotDmnd";
        //        case "horizontal brick":
        //            return  "horzBrick";
        //        case "large checker board":
        //            return  "lgCheck";
        //        case "large confetti":
        //            return  "lgConfetti";
        //        case "large grid":
        //            return  "lgGrid";
        //        case "light downward diagonal":
        //            return  "ltDnDiag";
        //        case "light horizontal":
        //            return  "ltHorz";
        //        case "light upward diagonal":
        //            return  "ltUpDiag";
        //        case "light vertical":
        //            return  "ltVert";
        //        case "narrow horizontal":
        //            return  "narHorz";
        //        case "narrow vertical":
        //            return  "narVert";
        //        case "outlined diamond":
        //            return  "openDmnd";
        //        case "small confetti":
        //            return  "smConfetti";
        //        case "small checker board":
        //            return  "smCheck";
        //        case "small grid":
        //            return  "smGrid";
        //        case "solid diamond":
        //            return  "solidDmnd";
        //        case "plaid":
        //            return  "plaid";
        //        case "shingle":
        //            return  "shingle";
        //        case "sphere":
        //            return  "sphere";
        //        case "trellis":
        //            return  "trellis";
        //        case "wave":
        //            return  "wave";
        //        case "weave":
        //            return  "weave";
        //        case "wide downward diagonal":
        //            return  "wdDnDiag";
        //        case "wide upward diagonal":
        //            return  "wdUpDiag";
        //        case "zig zag":
        //            return  "zigZag";
        //        default:
        //            return  "zigZag";
        //    }
        //}

        public static string getPrstForPatternCode(int code)
        {
            if (code == 0xC4) return "pct5";
            if (code == 0xC5) return "pct50";
            if (code == 0xC6) return "ltDnDiag";
            if (code == 0xC7) return "ltVert";
            if (code == 0xC8) return "dashDnDiag";
            if (code == 0xC9) return "zigZag";
            if (code == 0xCA) return "divot";
            if (code == 0xCB) return "smGrid";
            if (code == 0xCC) return "pct10";
            if (code == 0xCD) return "pct60";
            if (code == 0xCE) return "ltUpDiag";
            if (code == 0xCF) return "ltHorz";
            if (code == 0xD0) return "dashUpDiag";
            if (code == 0xD1) return "wave";
            if (code == 0xD2) return "dotGrid";
            if (code == 0xD3) return "lgGrid";
            if (code == 0xD4) return "pct20";
            if (code == 0xD5) return "pct70";
            if (code == 0xD6) return "dkDnDiag";
            if (code == 0xD7) return "narVert";
            if (code == 0xD8) return "dashHorz";
            if (code == 0xD9) return "diagBrick";
            if (code == 0xDA) return "dotDmnd";
            if (code == 0xDB) return "smCheck";
            if (code == 0xDC) return "pct25";
            if (code == 0xDD) return "pct75";
            if (code == 0xDE) return "dkUpDiag";
            if (code == 0xDF) return "narHorz";
            if (code == 0xE0) return "dashVert";
            if (code == 0xE1) return "horzBrick";
            if (code == 0xE2) return "shingle";
            if (code == 0xE3) return "lgCheck";
            if (code == 0xE4) return "pct30";
            if (code == 0xE5) return "pct80";
            if (code == 0xE6) return "wdDnDiag";
            if (code == 0xE7) return "dkVert";
            if (code == 0xE8) return "smConfetti";
            if (code == 0xE9) return "weave";
            if (code == 0xEA) return "trellis";
            if (code == 0xEB) return "openDmnd";
            if (code == 0xEC) return "pct40";
            if (code == 0xED) return "pct90";
            if (code == 0xEE) return "wdUpDiag";
            if (code == 0xEF) return "dkHorz";
            if (code == 0xF0) return "lgConfetti";
            if (code == 0xF1) return "plaid";
            if (code == 0xF2) return "sphere";
            if (code == 0xF3) return "solidDmnd";
            return "";
        }

        public static string getPrstForShape(uint shapeInstance)
        {
            switch (shapeInstance)
            {
                case 0x0: //NotPrimitive
                    return "";
                case 0x1: //Rectangle
                    return "rect";
                case 0x2: //RoundRectangle
                    return "roundRect";
                case 0x3: //ellipse
                    return "ellipse";
                case 0x4: //diamond
                    return "diamond";
                case 0x5: //triangle
                    return "triangle";
                case 0x6: //right triangle
                    return "rtTriangle";
                case 0x7: //parallelogram
                    return "parallelogram";
                case 0x8: //trapezoid
                    return "nonIsoscelesTrapezoid";
                case 0x9: //hexagon
                    return "hexagon";
                case 0xA: //octagon
                    return "octagon";
                case 0xB: //Plus
                    return "mathPlus";
                case 0xC: //Star
                    return "star5";
                case 0xD: //Arrow
                case 0xE: //ThickArrow
                    return "rightArrow";
                case 0xF: //HomePlate
                    return "homePlate";
                case 0x10: //Cube
                    return "cube";
                case 0x11: //Balloon
                    return "wedgeEllipseCallout";
                case 0x12: //Seal
                    return "star16";
                case 0x13: //Arc
                    return "curvedConnector2";
                case 0x14: //Line
                    return "line";
                case 0x15: //Plaque
                    return "plaque";
                case 0x16: //Cylinder
                    return "can";
                case 0x17: //Donut
                    return "donut";
                case 0x18: //TextSimple
                case 0x19: //TextOctagon
                case 0x1A: //TextHexagon
                case 0x1B: //TextCurve
                case 0x1C: //TextWave
                case 0x1D: //TextRing
                case 0x1E: //TextOnCurve
                case 0x1F: //TextOnRing
                    return "";
                case 0x20: //StraightConnector1
                    return "straightConnector1";
                case 0x21: //BentConnector2
                    return "bentConnector2";
                case 0x22: //BentConnector3
                    return "bentConnector3";
                case 0x23: //BentConnector4
                    return "bentConnector4";
                case 0x24: //BentConnector5
                    return "bentConnector5";
                case 0x25: //CurvedConnector2
                    return "curvedConnector2";
                case 0x26: //CurvedConnector3
                    return "curvedConnector3";
                case 0x27: //CurvedConnector4
                    return "curvedConnector4";
                case 0x28: //CurvedConnector5
                    return "curvedConnector5";
                case 0x29: //Callout1
                    return "callout1";
                case 0x2A: //Callout2
                    return "callout2";
                case 0x2B: //Callout3
                    return "callout3";
                case 0x2C: //AccentCallout1
                    return "accentCallout1";
                case 0x2D: //AccentCallout2
                    return "accentCallout2";
                case 0x2E: //AccentCallout3
                    return "accentCallout3";
                case 0x2F: //BorderCallout1
                    return "borderCallout1";
                case 0x30: //BorderCallout2
                    return "borderCallout2";
                case 0x31: //BorderCallout3
                    return "borderCallout3";
                case 0x32: //AccentBorderCallout1
                    return "accentBorderCallout1";
                case 0x33: //accentBorderCallout2
                    return "accentBorderCallout2";
                case 0x34: //accentBorderCallout3
                    return "accentBorderCallout3";
                case 0x35: //Ribbon
                    return "ribbon";
                case 0x36: //Ribbon2
                    return "ribbon2";
                case 0x37: //Chevron
                    return "chevron";
                case 0x38: //Pentagon
                    return "pentagon";
                case 0x39: //noSmoking
                    return "noSmoking";
                case 0x3A: //Seal8
                    return "star8";
                case 0x3B: //Seal16
                    return "star16";
                case 0x3C: //Seal32
                    return "star32";
                case 0x3D: //WedgeRectCallout
                    return "wedgeRectCallout";
                case 0x3E: //WedgeRRectCallout
                    return "wedgeRoundRectCallout";
                case 0x3F: //WedgeEllipseCallout
                    return "wedgeEllipseCallout";
                case 0x40: //Wave
                    return "wave";
                case 0x41: //FolderCorner
                    return "foldedCorner";
                case 0x42: //LeftArrow
                    return "leftArrow";
                case 0x43: //DownArrow
                    return "downArrow";
                case 0x44: //UpArrow
                    return "upArrow";
                case 0x45: //LeftRightArrow
                    return "leftRightArrow";
                case 0x46: //UpDownArrow
                    return "upDownArrow";
                case 0x47: //IrregularSeal1
                    return "irregularSeal1";
                case 0x48: //IrregularSeal2
                    return "irregularSeal2";
                case 0x49: //LightningBolt
                    return "lightningBolt";
                case 0x4A: //Heart
                    return "heart";
                case 0x4B: //PictureFrame
                    //return "frame";
                    return "rect";
                case 0x4C: //QuadArrow
                    return "quadArrow";
                case 0x4D: //LeftArrowCallout
                    return "leftArrowCallout";
                case 0x4E: //RightArrowCallout
                    return "rightArrowCallout";
                case 0x4F: //UpArrowCallout
                    return "upArrowCallout";
                case 0x50: //DownArrowCallout
                    return "downArrowCallout";
                case 0x51: //LeftRightArrowCallout
                    return "leftRightArrowCallout";
                case 0x52: //UpDownArrowCallout
                    return "upDownArrowCallout";
                case 0x53: //QuadArrowCallout
                    return "quadArrowCallout";
                case 0x54: //Bevel
                    return "bevel";
                case 0x55: //LeftBracket
                    return "leftBracket";
                case 0x56: //RightBracket
                    return "rightBracket";
                case 0x57: //LeftBrace
                    return "leftBrace";
                case 0x58: //RightBrace
                    return "rightBrace";
                case 0x59: //LeftUpArrow
                    return "leftUpArrow";
                case 0x5A: //BentUpArrow
                    return "bentUpArrow";
                case 0x5B: //BentArrow
                    return "bentArrow";
                case 0x5C: //Seal24
                    return "star24";
                case 0x5D: //stripedRightArrow
                    return "stripedRightArrow";
                case 0x5E: //notchedRightArrow
                    return "notchedRightArrow";
                case 0x5F: //BlockArc
                    return "blockArc";
                case 0x60: //SmileyFace
                    return "smileyFace";
                case 0x61: //verticalScroll
                    return "verticalScroll";
                case 0x62: //horizontalScroll
                    return "horizontalScroll";
                case 0x63: //circularArrow                        
                case 0x64: //notchedCircularArrow
                    return "circularArrow";
                case 0x65: //uturnArrow
                    return "uturnArrow";
                case 0x66: //curvedRightArrow
                    return "curvedRightArrow";
                case 0x67: //curvedLeftArrow
                    return "curvedLeftArrow";
                case 0x68: //curvedUpArrow
                    return "curvedUpArrow";
                case 0x69: //curvedDownArrow
                    return "curvedDownArrow";
                case 0x6A: //CloudCallout
                    return "cloudCallout";
                case 0x6B: //EllipseRibbon
                    return "ellipseRibbon";
                case 0x6C: //EllipseRibbon2
                    return "ellipseRibbon2";
                case 0x6D: //flowChartProcess
                    return "flowChartProcess";
                case 0x6E: //flowChartDecision
                    return "flowChartDecision";
                case 0x6F: //flowChartInputOutput
                    return "flowChartInputOutput";
                case 0x70: //flowChartPredefinedProcess
                    return "flowChartPredefinedProcess";
                case 0x71: //flowChartInternalStorage
                    return "flowChartInternalStorage";
                case 0x72: //flowChartDocument
                    return "flowChartDocument";
                case 0x73: //flowChartMultidocument
                    return "flowChartMultidocument";
                case 0x74: //flowChartTerminator
                    return "flowChartTerminator";
                case 0x75: //flowChartPreparation
                    return "flowChartPreparation";
                case 0x76: //flowChartManualInput
                    return "flowChartManualInput";
                case 0x77: //flowChartManualOperation
                    return "flowChartManualOperation";
                case 0x78: //flowChartConnector
                    return "flowChartConnector";
                case 0x79: //flowChartPunchedCard
                    return "flowChartPunchedCard";
                case 0x7A: //flowChartPunchedTape
                    return "flowChartPunchedTape";
                case 0x7B: //flowChartSummingJunction
                    return "flowChartSummingJunction";
                case 0x7C: //flowChartOr
                    return "flowChartOr";
                case 0x7D: //flowChartCollate
                    return "flowChartCollate";
                case 0x7E: //flowChartSort
                    return "flowChartSort";
                case 0x7F: //flowChartExtract
                    return "flowChartExtract";
                case 0x80: //flowChartMerge
                    return "flowChartMerge";
                case 0x81: //flowChartOfflineStorage
                    return "flowChartOfflineStorage";
                case 0x82: //flowChartOnlineStorage
                    return "flowChartOnlineStorage";
                case 0x83: //flowChartMagneticTape
                    return "flowChartMagneticTape";
                case 0x84: //flowChartMagneticDisk
                    return "flowChartMagneticDisk";
                case 0x85: //flowChartMagneticDrum
                    return "flowChartMagneticDrum";
                case 0x86: //flowChartDisplay
                    return "flowChartDisplay";
                case 0x87: //flowChartDelay
                    return "flowChartDelay";
                case 0x88: //TextPlainText
                case 0x89: //TextStop
                case 0x8A: //TextTriangle
                case 0x8B: //TextTriangleInverted
                case 0x8C: //TextChevron
                case 0x8D: //TextChevronInverted
                case 0x8E: //TextRingInside
                case 0x8F: //TextRingOutside
                case 0x90: //TextArchUpCurve
                case 0x91: //TextArchDownCurve
                case 0x92: //TextCircleCurve
                case 0x93: //TextButtonCurve
                case 0x94: //TextArchUpPour
                case 0x95: //TextArchDownPour
                case 0x96: //TextCirclePout
                case 0x97: //TextButtonPout
                case 0x98: //TextCurveUp
                case 0x99: //TextCurveDown
                case 0x9A: //TextCascadeUp
                case 0x9B: //TextCascadeDown
                case 0x9C: //TextWave1
                case 0x9D: //TextWave2
                case 0x9E: //TextWave3
                case 0x9F: //TextWave4
                case 0xA0: //TextInflate
                case 0xA1: //TextDeflate
                case 0xA2: //TextInflateBottom
                case 0xA3: //TextDeflateBottom
                case 0xA4: //TextInflateTop
                case 0xA5: //TextDeflateTop
                case 0xA6: //TextDeflateInflate
                case 0xA7: //TextDeflateInflateDeflate
                case 0xA8: //TextFadeRight
                case 0xA9: //TextFadeLeft
                case 0xAA: //TextFadeUp
                case 0xAB: //TextFadeDown
                case 0xAC: //TextSlantUp
                case 0xAD: //TextSlantDown
                case 0xAE: //TextCanUp
                case 0xAF: //TextCanDown
                    return "";
                case 0xB0: //flowchartAlternateProcess
                    return "flowChartAlternateProcess";
                case 0xB1: //flowchartOffpageConnector
                    return "flowChartOffpageConnector";
                case 0xB2: //Callout90
                    return "callout1";
                case 0xB3: //AccentCallout90
                    return "accentCallout1";
                case 0xB4: //BorderCallout90
                    return "borderCallout1";
                case 0xB5: //AccentBorderCallout90
                    return "accentBorderCallout1";
                case 0xB6: //LeftRightUpArrow
                    return "leftRightUpArrow";
                case 0xB7: //Sun
                    return "sun";
                case 0xB8: //Moon
                    return "moon";
                case 0xB9: //BracketPair
                    return "bracketPair";
                case 0xBA: //BracePair
                    return "bracePair";
                case 0xBB: //Seal4
                    return "star4";
                case 0xBC: //DoubleWave
                    return "doubleWave";
                case 0xBD: //ActionButtonBlank
                    return "actionButtonBlank";
                case 0xBE: //ActionButtonHome
                    return "actionButtonHome";
                case 0xBF: //ActionButtonHelp
                    return "actionButtonHelp";
                case 0xC0: //ActionButtonInformation
                    return "actionButtonInformation";
                case 0xC1: //ActionButtonForwardNext
                    return "actionButtonForwardNext";
                case 0xC2: //ActionButtonBackPrevious
                    return "actionButtonBackPrevious";
                case 0xC3: //ActionButtonEnd
                    return "actionButtonEnd";
                case 0xC4: //ActionButtonBeginning
                    return "actionButtonBeginning";
                case 0xC5: //ActionButtonReturn
                    return "actionButtonReturn";
                case 0xC6: //ActionButtonDocument
                    return "actionButtonDocument";
                case 0xC7: //ActionButtonSound
                    return "actionButtonSound";
                case 0xC8: //ActionButtonMovie
                    return "actionButtonMovie";
                case 0xC9: //HostControl (do not use)
                    return "";
                case 0xCA: //TextBox
                    return "rect";
                default:
                    return "";
            }
        }
    }
}
