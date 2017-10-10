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
using System.Reflection;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public abstract class BiffRecord
    {
        IStreamReader _reader;

        RecordType _id;
        uint _length;
        long _offset;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Record ID - Recordtype</param>
        /// <param name="length">The recordlegth</param>
        public BiffRecord(IStreamReader reader, RecordType id, ushort length)
        {
            _reader = reader;
            _offset = _reader.BaseStream.Position;

            _id = id;
            _length = length;
        }

        private static Dictionary<ushort, Type> TypeToRecordClassMapping = new Dictionary<ushort, Type>();

        static BiffRecord()
        {
            UpdateTypeToRecordClassMapping(
                Assembly.GetExecutingAssembly(),
                typeof(BOF).Namespace);
        }

        public static void UpdateTypeToRecordClassMapping(Assembly assembly, string ns)
        {
            foreach (var t in assembly.GetTypes())
            {
                if (ns == null || t.Namespace == ns)
                {
                    var attrs = t.GetCustomAttributes(typeof(BiffRecordAttribute), false);

                    BiffRecordAttribute attr = null;

                    if (attrs.Length > 0)
                        attr = attrs[0] as BiffRecordAttribute;

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


        public static RecordType GetNextRecordType(IStreamReader reader)
        {
            long position = reader.BaseStream.Position;
                
            // read type of the next record
            var nextRecord = (RecordType)reader.ReadUInt16();
            var length = reader.ReadUInt16();

            // skip leading StartBlock/EndBlock records
            if (nextRecord == RecordType.StartBlock
                || nextRecord == RecordType.EndBlock 
                || nextRecord == RecordType.StartObject
                || nextRecord == RecordType.EndObject
                || nextRecord == RecordType.ChartFrtInfo)
            {
                // skip the body of the record
                reader.ReadBytes(length);
                // get the type of the next record
                return GetNextRecordType(reader);
            }
            else if (nextRecord == RecordType.FrtWrapper)
            {
                // return type of wrapped Biff record
                var frtWrapper = new FrtWrapper(reader, nextRecord, length);
                reader.BaseStream.Position = position;
                return frtWrapper.wrappedRecord.Id;
            }
            else
            {
                // seek back to the begin of the current record
                reader.BaseStream.Position = position;
                return nextRecord;
            }
        }

        public static BiffRecord ReadRecord(IStreamReader reader)
        {
            BiffRecord result = null;
            try
            {
                var id = (RecordType)reader.ReadUInt16();
                var length = reader.ReadUInt16();

                // skip leading StartBlock/EndBlock records
                if (id == RecordType.StartBlock ||
                    id == RecordType.EndBlock ||
                    id == RecordType.StartObject ||
                    id == RecordType.EndObject ||
                    id == RecordType.ChartFrtInfo)
                {
                    // skip the body of this record
                    reader.ReadBytes(length);

                    // get the next record
                    return ReadRecord(reader);
                }
                else if (id == RecordType.FrtWrapper)
                {
                    // return type of wrapped Biff record
                    var frtWrapper = new FrtWrapper(reader, id, length);
                    return frtWrapper.wrappedRecord;
                }

                Type cls;
                if (TypeToRecordClassMapping.TryGetValue((ushort)id, out cls))
                {
                    var constructor = cls.GetConstructor(
                        new Type[] { typeof(IStreamReader), typeof(RecordType), typeof(ushort) }
                        );

                    try
                    {
                        result = (BiffRecord)constructor.Invoke(
                            new object[] { reader, id, length }
                            );
                    }
                    catch (TargetInvocationException e)
                    {
                        throw e.InnerException;
                    }
                }
                else
                {
                    result = new UnknownBiffRecord(reader, (RecordType)id, length);
                }

                return result;
            }
            catch (OutOfMemoryException e)
            {
                throw new Exception("Invalid BIFF record", e);
            }
        }

        public RecordType Id
        {
            get { return _id; }
        }

        public uint Length
        {
            get { return _length; }
        }

        public long Offset
        {
            get { return _offset; }
        }

        public IStreamReader Reader
        {
            get { return _reader; }
            set { this._reader = value; }
        }


    }
}
