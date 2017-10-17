
using System;
using System.Collections.Generic;
using System.Reflection;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
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
            this._reader = reader;
            this._offset = this._reader.BaseStream.Position;

            this._id = id;
            this._length = length;
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
                                throw new Exception(string.Format(
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
            ushort length = reader.ReadUInt16();

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
                ushort length = reader.ReadUInt16();

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
            get { return this._id; }
        }

        public uint Length
        {
            get { return this._length; }
        }

        public long Offset
        {
            get { return this._offset; }
        }

        public IStreamReader Reader
        {
            get { return this._reader; }
            set { this._reader = value; }
        }


    }
}
