

using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public abstract class AbstractChartGroupMapping : AbstractChartMapping,
          IMapping<CrtSequence>
    {
        protected bool _is3DChart;

        public AbstractChartGroupMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext)
        {
            this._is3DChart = is3DChart;
        }

        public bool Is3DChart
        {
            get { return this._is3DChart; }
        }

        #region IMapping<CrtSequence> Members

        public abstract void Apply(CrtSequence crtSequence);

        #endregion
    }
}
