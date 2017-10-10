/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using System;
using System.Collections;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkChain
    {
        public byte recordVersion;

        public ushort xmltkParent;

        public ArrayList chainRecords = new ArrayList();       

        public XmlTkChain(IStreamReader reader)
        {
            this.recordVersion = reader.ReadByte();

            //unused
            reader.ReadByte();

            this.xmltkParent = reader.ReadUInt16();

            long pos;
            XmlTkHeader h;
            switch (this.xmltkParent)
            {
                case 0x01:
                    //chainRecords = [XmlTkMaxFrt] [XmlTkMinFrt] [XmlTkLogBaseFrt]

                    if (getNextXmlTkTag(reader) == 0x55)
                    {
                        this.chainRecords.Add(new XmlTkMaxFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x56)
                    {
                        this.chainRecords.Add(new XmlTkMinFrt(reader));
                    }

                     if (getNextXmlTkTag(reader) == 0x0)
                    {
                         this.chainRecords.Add(new XmlTkLogBaseFrt(reader));
                    }

                    break;
                case 0x02:
                    //chainRecords = [XmlTkStyle] [XmlTkThemeOverride] [XmlTkColorMappingOverride]
                    
                    if(getNextXmlTkTag(reader) == 0x03)
                    {
                        this.chainRecords.Add(new XmlTkStyle(reader));
                    }

                    if(getNextXmlTkTag(reader) == 0x33)
                    {
                        this.chainRecords.Add(new XmlTkThemeOverride(reader));
                    }

                    if(getNextXmlTkTag(reader) == 0x34)
                    {
                        this.chainRecords.Add(new XmlTkColorMappingOverride(reader));
                    }

                    break;
                case 0x04:
                    //chainRecords = [XmlTkNoMultiLvlLbl] [XmlTkTickLabelSkipFrt] [XmlTkTickMarkSkipFrt] [XmlTkMajorUnitFrt] 
                    //[XmlTkMinorUnitFrt] [XmlTkTickLabelPositionFrt] [XmlTkBaseTimeUnitFrt] [XmlTkFormatCodeFrt] [XmlTkMajorUnitTypeFrt] [XmlTkMinorUnitTypeFrt]

                    if (getNextXmlTkTag(reader) == 0x2E)
                    {
                        this.chainRecords.Add(new XmlTkNoMultiLvlLbl(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x51)
                    {
                        this.chainRecords.Add(new XmlTkTickLabelSkipFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x52)
                    {
                        this.chainRecords.Add(new XmlTkTickMarkSkipFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x53)
                    {
                        this.chainRecords.Add(new XmlTkMajorUnitFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x54)
                    {
                        this.chainRecords.Add(new XmlTkMinorUnitFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x5C)
                    {
                        this.chainRecords.Add(new XmlTkTickLabelPositionFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x5F)
                    {
                        this.chainRecords.Add(new XmlTkBaseTimeUnitFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x64)
                    {
                        this.chainRecords.Add(new XmlTkFormatCodeFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x6A)
                    {
                        this.chainRecords.Add(new XmlTkMajorUnitTypeFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x6B)
                    {
                        this.chainRecords.Add(new XmlTkMinorUnitTypeFrt(reader));
                    }
                    
                    break;
                case 0x05:
                    //chainRecords = [XmlTkShowDLblsOverMax] [XmlTkBackWallThicknessFrt] [XmlTkFloorThicknessFrt] [XmlTkDispBlanksAsFrt] [SURFACE]
                    //SURFACE = XmlTkStartSurface [XmlTkFormatCodeFrt [XmlTkSpb]] [XmlTkTpb] XmlTkEndSurface

                    bool first = true;

                    if (getNextXmlTkTag(reader) == 0x5B)
                    {
                        this.chainRecords.Add(new XmlTkShowDLblsOverMax(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x35)
                    {
                        this.chainRecords.Add(new XmlTkBackWallThicknessFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x36)
                    {
                        this.chainRecords.Add(new XmlTkFloorThicknessFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x66)
                    {
                        this.chainRecords.Add(new XmlTkDispBlanksAsFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x59)
                    {
                        if (first)
                        {
                            this.chainRecords.Add(new XmlTkStartSurface(reader));
                            first = false;
                        }
                        else
                        {
                            this.chainRecords.Add(new XmlTkEndSurface(reader));
                        }
                    }

                    if (getNextXmlTkTag(reader) == 0x64)
                    {
                        this.chainRecords.Add(new XmlTkFormatCodeFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x1E)
                    {
                        this.chainRecords.Add(new XmlTkSpb(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x20)
                    {
                        this.chainRecords.Add(new XmlTkTpb(reader));
                    }
                    break;
                case 0x0F:
                    //chainRecords = [XmlTkOverlay]
                    this.chainRecords.Add(new XmlTkOverlay(reader));
                    break;
                case 0x13:
                    //chainRecords = [XmlTkSymbolFrt]
                    this.chainRecords.Add(new XmlTkSymbolFrt(reader));
                    break;
                case 0x16:
                    //chainRecords = [XmlTkPieComboFrom12Frt]
                    this.chainRecords.Add(new XmlTkPieComboFrom12Frt(reader));
                    break;
                case 0x19:
                    //chainRecords = [XmlTkOverlay]
                    this.chainRecords.Add(new XmlTkOverlay(reader));
                    break;
                case 0x37:
                    //chainRecords = [XmlTkRAngAxOffFrt] [XmlTkPerspectiveFrt] [XmlTkRotYFrt] [XmlTkRotXFrt] [XmlTkHeightPercent]

                    if (getNextXmlTkTag(reader) == 0x50)
                    {
                        this.chainRecords.Add(new XmlTkRAngAxOffFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x4D)
                    {
                        this.chainRecords.Add(new XmlTkPerspectiveFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x4F)
                    {
                        this.chainRecords.Add(new XmlTkRotYFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x4E)
                    {
                        this.chainRecords.Add(new XmlTkRotXFrt(reader));
                    }

                    if (getNextXmlTkTag(reader) == 0x65)
                    {
                        this.chainRecords.Add(new XmlTkHeightPercent(reader));
                    }
                    
                    break;
            }

            // ignore remaing record data
            //reader.ReadBytes(8);
        }

        public ushort getNextXmlTkTag(IStreamReader reader)
        {
            long pos = reader.BaseStream.Position;
            var header = new XmlTkHeader(reader);
            reader.BaseStream.Position = pos;

            return header.xmlTkTag;
        }
    }
}
