using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class AxesSequence : BiffRecordSequence, IVisitable
    {
        public IvAxisSequence IvAxisSequence;

        public DvAxisSequence DvAxisSequence;

        public DvAxisSequence DvAxisSequence2;

        // NOTE: SeriesAxisSequence is just a simplified IvAxisSequence
        //public SeriesAxisSequence SeriesAxisSequence;
        public IvAxisSequence SeriesAxisSequence;

        public List<AttachedLabelSequence> AttachedLabelSequences;

        public PlotArea PlotArea;

        public FrameSequence Frame;

        public AxesSequence(IStreamReader reader)
            : base(reader)
        {
            //AXES = [IVAXIS DVAXIS [SERIESAXIS] / DVAXIS DVAXIS] *3ATTACHEDLABEL [PlotArea FRAME]

            // [IVAXIS DVAXIS [SERIESAXIS] / DVAXIS DVAXIS]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Axis)
            {
                long position = reader.BaseStream.Position;

                Axis axis = (Axis)BiffRecord.ReadRecord(reader);

                Begin begin = (Begin)BiffRecord.ReadRecord(reader);

                if (BiffRecord.GetNextRecordType(reader) == RecordType.CatSerRange
                    || BiffRecord.GetNextRecordType(reader) == RecordType.AxcExt)
                {
                    reader.BaseStream.Position = position;

                    ChartAxisIdGenerator.Instance.StartNewAxisGroup();

                    //IVAXIS
                    this.IvAxisSequence = new IvAxisSequence(reader);

                    //DVAXIS 
                    this.DvAxisSequence = new DvAxisSequence(reader);
                    
                    //[SERIESAXIS]  
                    if (BiffRecord.GetNextRecordType(reader) == RecordType.Axis)
                    {
                        this.SeriesAxisSequence = new IvAxisSequence(reader);
                    }
                }
                else
                {
                    reader.BaseStream.Position = position;

                    ChartAxisIdGenerator.Instance.StartNewAxisGroup();

                    //DVAXIS 
                    this.DvAxisSequence = new DvAxisSequence(reader);
                    
                    //DVAXIS 
                    this.DvAxisSequence2 = new DvAxisSequence(reader);
                }
            }

            //*3ATTACHEDLABEL 
            this.AttachedLabelSequences = new List<AttachedLabelSequence>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Text)
            {
                this.AttachedLabelSequences.Add(new AttachedLabelSequence(reader));
            }

            //[PlotArea FRAME]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.PlotArea)
            {
                this.PlotArea = (PlotArea)BiffRecord.ReadRecord(reader);

                this.Frame = new FrameSequence(reader);
            }
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<AxesSequence>)mapping).Apply(this);
        }

        #endregion
    }
}
