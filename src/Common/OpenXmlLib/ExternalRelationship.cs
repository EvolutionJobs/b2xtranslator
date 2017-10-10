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

namespace DIaLOGIKa.b2xtranslator.OpenXmlLib
{
    public class ExternalRelationship
    {
        protected string _id;
        protected string _relationshipType;
        protected string _target;
        
        public ExternalRelationship(string id, string relationshipType, Uri targetUri)
        {
            _id = id;
            _relationshipType = relationshipType;
            _target = targetUri.ToString();
        }

        public ExternalRelationship(string id, string relationshipType, string target)
        {
            _id = id;
            _relationshipType = relationshipType;
            _target = target;
        }

        public string Id
        {
            get { return _id; }
            set { _id = value; }
        }
        
        public string RelationshipType
        {
            get { return _relationshipType; }
            set { _relationshipType = value; }
        }

        public string Target
        {
            get { return _target; }
            set { _target = value; }
        }

        public Uri TargetUri
        {
            get { return new Uri(_target, UriKind.RelativeOrAbsolute); }
        }
    }
}
