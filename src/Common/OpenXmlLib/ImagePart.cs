/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib
{
    public class ImagePart : OpenXmlPart
    {
        public enum ImageType
        {
            Bmp,
            Emf,
            Gif,
            Icon,
            Jpeg,
            //Pcx,
            Png,
            Tiff,
            Wmf
        }

        protected ImageType _type;

        internal ImagePart(ImageType type, OpenXmlPartContainer parent, int partIndex)
            : base(parent, partIndex)
        {
            _type = type;
        }

        public override string ContentType
        {
            get 
            {
                switch (_type)
                {
                    case ImageType.Bmp:
                        return "image/bmp";
                    case ImageType.Emf:
                        return "image/x-emf";
                    case ImageType.Gif:
                        return "image/gif";
                    case ImageType.Icon:
                        return "image/x-icon";
                    case ImageType.Jpeg:
                        return "image/jpeg";
                    //case ImagePartType.Pcx:
                    //    return "image/pcx";
                    case ImageType.Png:
                        return "image/png";
                    case ImageType.Tiff:
                        return "image/tiff";
                    case ImageType.Wmf:
                        return "image/x-wmf";
                    default:
                        return "image/png";
                }
            }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Image; }
        }

        public override string TargetName { get { return "image" + this.PartIndex; } }
        //public override string TargetDirectory { get { return "media"; } }

        private string targetdirectory = "media";
        public override string TargetDirectory
        {
            get
            {
                return targetdirectory;
            }

            set
            {
                targetdirectory = value;
            }

        }

        public override string TargetExt
        {
            get
            {
                switch (_type)
                {
                    case ImageType.Bmp:
                        return ".bmp";
                    case ImageType.Emf:
                        return ".emf";
                    case ImageType.Gif:
                        return ".gif";
                    case ImageType.Icon:
                        return ".ico";
                    case ImageType.Jpeg:
                        return ".jpg";
                    //case ImagePartType.Pcx:
                    //    return ".pcx";
                    case ImageType.Png:
                        return ".png";
                    case ImageType.Tiff:
                        return ".tif";
                    case ImageType.Wmf:
                        return ".wmf";
                    default:
                        return ".png";
                }
            }
        }
    }
}
