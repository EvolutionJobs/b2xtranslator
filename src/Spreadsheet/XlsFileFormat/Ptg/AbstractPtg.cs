using System;
using System.Globalization;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public abstract class AbstractPtg
    {
        IStreamReader _reader;
        PtgNumber _id;
        long _offset;
        string data;
        uint length;
        
        protected uint popSize;
        protected PtgType type; 


        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Ptg Id</param>
        /// <param name="length">The recordlength</param>
        public AbstractPtg(IStreamReader reader, PtgNumber ptgid)
        {
            this._reader = reader;
            this._offset = this._reader.BaseStream.Position;
            this._id = ptgid;
            this.data = ""; 
        }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Ptg Id</param>
        /// <param name="length">The recordlength</param>
        public AbstractPtg(IStreamReader reader, Ptg0x18Sub ptgid)
        {
            this._reader = reader;
            this._offset = this._reader.BaseStream.Position;
            this._id = (PtgNumber)ptgid;
            this.data = "";
        }

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Ptg Id</param>
        /// <param name="length">The recordlength</param>
        public AbstractPtg(IStreamReader reader, Ptg0x19Sub ptgid)
        {
            this._reader = reader;
            this._offset = this._reader.BaseStream.Position;
            this._id = (PtgNumber)ptgid;
            this.data = "";
        }

        public PtgNumber Id
        {
            get { return this._id; }
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

        protected string Data
        {
            get { return this.data; }
            set { this.data = value; }
        }

        protected uint Length
        {
            get { return this.length; }
            set { this.length = value; }
        }

        public uint getLength()
        {
            return this.length; 
        }

        public string getData()
        {            
            return Convert.ToString(this.data,CultureInfo.GetCultureInfo("en-US"));
        }

        public uint PopSize()
        {
            return this.popSize;
        }

        public PtgType OpType()
        {
            return this.type; 
        }

    }
}
