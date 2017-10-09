using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.DrawingML;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class DataPointMapping : AbstractChartMapping,
          IMapping<SsSequence>
    {
        int _index;

        public DataPointMapping(ExcelContext workbookContext, ChartContext chartContext, int index)
            : base(workbookContext, chartContext)
        {
            _index = index;
        }

        public void Apply(SsSequence ssSequence)
        {
            // c:dPt
            _writer.WriteStartElement(Dml.Chart.Prefix, "dPt", Dml.Chart.Ns);
            {
                // c:bubble3D
                // TODO

                // c:explosion
                // TODO

                // c:idx
                writeValueElement("idx", _index.ToString());

                // c:invertIfNegative
                // TODO

                // c:marker
                // TODO

                // c:pictureOptions
                // TODO

                // c:spPr
                ssSequence.Convert(new ShapePropertiesMapping(this.WorkbookContext, this.ChartContext));
            }
            _writer.WriteEndElement();
        }
    }
}
