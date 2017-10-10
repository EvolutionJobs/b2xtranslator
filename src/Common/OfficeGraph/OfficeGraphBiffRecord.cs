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
using System.Collections.Generic;
using System.Reflection;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.OfficeGraph
{
    public abstract class OfficeGraphBiffRecord
    {
        GraphRecordNumber _id;
        UInt32 _length;
        long _offset;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Record ID - Recordtype</param>
        /// <param name="length">The recordlegth</param>
        public OfficeGraphBiffRecord(IStreamReader reader, GraphRecordNumber id, UInt32 length)
        {
            _reader = reader;
            _offset = _reader.BaseStream.Position;

            _id = id;
            _length = length;
        }

        private static Dictionary<ushort, Type> TypeToRecordClassMapping = new Dictionary<ushort, Type>();

        static OfficeGraphBiffRecord()
        {
            UpdateTypeToRecordClassMapping(
                Assembly.GetExecutingAssembly(), 
                typeof(OfficeGraphBiffRecord).Namespace);
        }

        public static void UpdateTypeToRecordClassMapping(Assembly assembly, string ns)
        {
            foreach (var t in assembly.GetTypes())
            {
                if (ns == null || t.Namespace == ns)
                {
                    var attrs = t.GetCustomAttributes(typeof(OfficeGraphBiffRecordAttribute), false);

                    OfficeGraphBiffRecordAttribute attr = null;

                    if (attrs.Length > 0)
                        attr = attrs[0] as OfficeGraphBiffRecordAttribute;

                    if (attr != null)
                    {
                        // Add the type codes of the array
                        foreach (ushort typeCode in attr.TypeCodes)
                        {
                            if (TypeToRecordClassMapping.ContainsKey(typeCode))
                            {
                                throw new Exception(String.Format(
                                    "Tried to register TypeCode {0} to {1}, but it is already registered to {2}",
                                    typeCode, t, TypeToRecordClassMapping[typeCode]));
                            }
                            TypeToRecordClassMapping.Add(typeCode, t);
                        }
                    }
                }
            }
        }

        [Obsolete("Use OfficeGraphBiffRecordSequence.GetNextRecordNumber")]
        public static GraphRecordNumber GetNextRecordNumber(IStreamReader reader)
        {
            // read next id
            var nextRecord = (GraphRecordNumber)reader.ReadUInt16();

            // seek back
            reader.BaseStream.Seek(-sizeof(ushort), System.IO.SeekOrigin.Current);

            return nextRecord;
        }

        [Obsolete("Use OfficeGraphBiffRecordSequence.ReadRecord")]
        public static OfficeGraphBiffRecord ReadRecord(IStreamReader reader)
        {
            OfficeGraphBiffRecord result = null;
            try
            {
                var id = reader.ReadUInt16();
                var size = reader.ReadUInt16();
                Type cls;
                if (TypeToRecordClassMapping.TryGetValue(id, out cls))
                {
                    var constructor = cls.GetConstructor(
                        new Type[] { typeof(IStreamReader), typeof(GraphRecordNumber), typeof(ushort) }
                        );

                    try
                    {
                        result = (OfficeGraphBiffRecord)constructor.Invoke(
                            new object[] {reader, id, size }
                            );
                    }
                    catch (TargetInvocationException e)
                    {
                        throw e.InnerException;
                    }
                }
                else
                {
                    result = new UnknownGraphRecord(reader, id, size);
                }

                return result;
            }
            catch (OutOfMemoryException e)
            {
                throw new Exception("Invalid record", e);
            }
        }

        public GraphRecordNumber Id
        {
            get { return _id; }
        }

        public UInt32 Length
        {
            get { return _length; }
        }

        public long Offset
        {
            get { return _offset; }
        }

        IStreamReader _reader;
        public IStreamReader Reader
        {
            get { return _reader; }
            set { this._reader = value; }
        }
    }
}
