/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
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

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecordAttribute(0xF006)]
    public class DrawingGroupRecord : Record
    {
        public class FileIdCluster
        {
            public UInt32 DrawingGroupId;
            public UInt32 CSpIdCur;

            public FileIdCluster(BinaryReader reader)
            {
                this.DrawingGroupId = reader.ReadUInt32();
                this.CSpIdCur = reader.ReadUInt32();
            }

            public string ToString(uint depth)
            {
                StringBuilder result = new StringBuilder();

                result.Append(IndentationForDepth(depth));
                result.AppendFormat("FileIdCluster: DrawingGroupId = {0}, CSpIdCur = {1}",
                   this.DrawingGroupId, this.CSpIdCur);

                return result.ToString();
            }
        }

        public UInt32 MaxShapeId;           // Maximum shape ID
        public UInt32 IdClustersCount;      // Number of FileIdClusters
        public UInt32 ShapesSavedCount;     // Total number of shapes saved
        public UInt32 DrawingsSavedCount;   // Total number of drawings saved

        public List<FileIdCluster> Clusters = new List<FileIdCluster>();

        public DrawingGroupRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.MaxShapeId = this.Reader.ReadUInt32();
            this.IdClustersCount = this.Reader.ReadUInt32() - 1; // Office saves the actual value + 1 -- flgr
            this.ShapesSavedCount = this.Reader.ReadUInt32();
            this.DrawingsSavedCount = this.Reader.ReadUInt32();

            for (int i = 0; i < this.IdClustersCount; i++)
            {
                Clusters.Add(new FileIdCluster(this.Reader));
            }
        }

        override public string ToString(uint depth)
        {
            StringBuilder result = new StringBuilder();

            result.AppendLine(base.ToString(depth));

            result.Append(IndentationForDepth(depth + 1));
            result.AppendFormat("MaxShapeId = {0}, IdClustersCount = {1}",
                this.MaxShapeId, this.IdClustersCount);

            result.AppendLine();
            result.Append(IndentationForDepth(depth + 1));
            result.AppendFormat("ShapesSavedCount = {0}, DrawingsSavedCount = {1}",
                this.ShapesSavedCount, this.DrawingsSavedCount);

            depth++;

            if (this.Clusters.Count > 0)
            {
                result.AppendLine();
                result.Append(IndentationForDepth(depth));
                result.Append("Clusters:");
            }

            foreach (FileIdCluster cluster in this.Clusters)
            {
                result.AppendLine();
                result.Append(cluster.ToString(depth + 1));
            }

            return result.ToString();
        }
    }

}
