using System;
using System.IO;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// This class should extract the specific worksheet data. 
    /// </summary>
    public class WorksheetExtractor : Extractor
    {
        /// <summary>
        /// Datacontainer for the worksheet
        /// </summary>
        private WorkSheetData bsd;

        /// <summary>
        /// CTor 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="bsd"> Boundsheetdata container</param>
        public WorksheetExtractor(VirtualStreamReader reader, WorkSheetData bsd)
            : base(reader) 
        {
            this.bsd = bsd;
            this.extractData(); 
        }

        /// <summary>
        /// Extracting the data from the stream 
        /// </summary>
        public override void extractData()
        {
            BiffHeader bh, latestbiff;
            BOF firstBOF = null;
            
            //try
            //{
                while (this.StreamReader.BaseStream.Position < this.StreamReader.BaseStream.Length)
                {
                    bh.id = (RecordType)this.StreamReader.ReadUInt16();
                    bh.length = this.StreamReader.ReadUInt16();

                    TraceLogger.DebugInternal("BIFF {0}\t{1}\t", bh.id, bh.length);

                    switch (bh.id)
                    {
                        case RecordType.EOF:
                            {
                                this.StreamReader.BaseStream.Seek(0, SeekOrigin.End);
                            }
                            break;
                        case RecordType.BOF:
                            {
                                BOF bof = new BOF(this.StreamReader, bh.id, bh.length);

                                switch (bof.docType)
                                {
                                    case BOF.DocumentType.WorkbookGlobals:
                                    case BOF.DocumentType.Worksheet:
                                        firstBOF = bof;
                                        break;

                                    case BOF.DocumentType.Chart:
                                        // parse chart 

                                        break;

                                    default:
                                        this.readUnkownFile();
                                        break;
                                }
                            }
                            break;
                        case RecordType.LabelSst:
                            {
                                LabelSst labelsst = new LabelSst(this.StreamReader, bh.id, bh.length);
                                this.bsd.addLabelSST(labelsst);
                            }
                            break;
                        case RecordType.MulRk:
                            {
                                MulRk mulrk = new MulRk(this.StreamReader, bh.id, bh.length);
                                this.bsd.addMULRK(mulrk);
                            }
                            break;
                        case RecordType.Number:
                            {
                                Number number = new Number(this.StreamReader, bh.id, bh.length);
                                this.bsd.addNUMBER(number);
                            }
                            break;
                        case RecordType.RK:
                            {
                                RK rk = new RK(this.StreamReader, bh.id, bh.length);
                                this.bsd.addRK(rk);
                            }
                            break;
                        case RecordType.MergeCells:
                            {
                                MergeCells mergecells = new MergeCells(this.StreamReader, bh.id, bh.length);
                                this.bsd.MERGECELLSData = mergecells;
                            }
                            break;
                        case RecordType.Blank:
                            {
                                Blank blankcell = new Blank(this.StreamReader, bh.id, bh.length);
                                this.bsd.addBLANK(blankcell);
                            } break;
                        case RecordType.MulBlank:
                            {
                                MulBlank mulblank = new MulBlank(this.StreamReader, bh.id, bh.length);
                                this.bsd.addMULBLANK(mulblank);
                            }
                            break;
                        case RecordType.Formula:
                            {
                                Formula formula = new Formula(this.StreamReader, bh.id, bh.length);
                                this.bsd.addFORMULA(formula);
                                TraceLogger.DebugInternal(formula.ToString());
                            }
                            break;
                        case RecordType.Array:
                            {
                                ARRAY array = new ARRAY(this.StreamReader, bh.id, bh.length);
                                this.bsd.addARRAY(array);
                            }
                            break;
                        case RecordType.ShrFmla:
                            {
                                ShrFmla shrfmla = new ShrFmla(this.StreamReader, bh.id, bh.length);
                                this.bsd.addSharedFormula(shrfmla);

                            }
                            break;
                        case RecordType.String:
                            {
                                STRING formulaString = new STRING(this.StreamReader, bh.id, bh.length);
                                this.bsd.addFormulaString(formulaString.value);

                            }
                            break;
                        case RecordType.Row:
                            {
                                Row row = new Row(this.StreamReader, bh.id, bh.length);
                                this.bsd.addRowData(row);

                            }
                            break;
                        case RecordType.ColInfo:
                            {
                                ColInfo colinfo = new ColInfo(this.StreamReader, bh.id, bh.length);
                                this.bsd.addColData(colinfo);
                            }
                            break;
                        case RecordType.DefColWidth:
                            {
                                DefColWidth defcolwidth = new DefColWidth(this.StreamReader, bh.id, bh.length);
                                this.bsd.addDefaultColWidth(defcolwidth.cchdefColWidth);
                            }
                            break;
                        case RecordType.DefaultRowHeight:
                            {
                                DefaultRowHeight defrowheigth = new DefaultRowHeight(this.StreamReader, bh.id, bh.length);
                                this.bsd.addDefaultRowData(defrowheigth);
                            }
                            break;
                        case RecordType.LeftMargin:
                            {
                                LeftMargin leftm = new LeftMargin(this.StreamReader, bh.id, bh.length);
                                this.bsd.leftMargin = leftm.value;
                            }
                            break;
                        case RecordType.RightMargin:
                            {
                                RightMargin rightm = new RightMargin(this.StreamReader, bh.id, bh.length);
                                this.bsd.rightMargin = rightm.value;
                            }
                            break;
                        case RecordType.TopMargin:
                            {
                                TopMargin topm = new TopMargin(this.StreamReader, bh.id, bh.length);
                                this.bsd.topMargin = topm.value;
                            }
                            break;
                        case RecordType.BottomMargin:
                            {
                                BottomMargin bottomm = new BottomMargin(this.StreamReader, bh.id, bh.length);
                                this.bsd.bottomMargin = bottomm.value;
                            }
                            break;
                        case RecordType.Setup:
                            {
                                Setup setup = new Setup(this.StreamReader, bh.id, bh.length);
                                this.bsd.addSetupData(setup);
                            }
                            break;
                        case RecordType.HLink:
                            {
                                long oldStreamPos = this.StreamReader.BaseStream.Position;
                                try
                                {
                                    HLink hlink = new HLink(this.StreamReader, bh.id, bh.length);
                                    bsd.addHyperLinkData(hlink);
                                }
                                catch (Exception ex)
                                {
                                    this.StreamReader.BaseStream.Seek(oldStreamPos, System.IO.SeekOrigin.Begin);
                                    this.StreamReader.BaseStream.Seek(bh.length, System.IO.SeekOrigin.Current);
                                    TraceLogger.Debug("Link parse error");
                                    TraceLogger.Error(ex.StackTrace);
                                }
                            }
                            break;
                        case RecordType.MsoDrawing:
                            {
                                // Record header has already been read. Reset position to record beginning.
                                this.StreamReader.BaseStream.Position -= 2 * sizeof(UInt16);
                                this.bsd.ObjectsSequence = new ObjectsSequence(this.StreamReader);
                            }
                            break;
                        default:
                            {
                                // this else statement is used to read BiffRecords which aren't implemented 
                                byte[] buffer = new byte[bh.length];
                                buffer = this.StreamReader.ReadBytes(bh.length);
                            }
                            break;
                    }
                    latestbiff = bh; 
                }
            //}
            //catch (Exception ex)
            //{
            //    TraceLogger.Error(ex.Message);
            //    TraceLogger.Error(ex.StackTrace); 
            //    TraceLogger.Debug(ex.ToString());
            //}
        }

        /// <summary>
        /// This method should read over every record which is inside a file in the worksheet file 
        /// For example this could be the diagram "file" 
        /// A diagram begins with the BOF Biffrecord and ends with the EOF record. 
        /// </summary>
        public void readUnkownFile(){
            BiffHeader bh;
            //try
            //{
                do
                {
                    bh.id = (RecordType)this.StreamReader.ReadUInt16();
                    bh.length = this.StreamReader.ReadUInt16();
                    this.StreamReader.ReadBytes(bh.length);
                } while (bh.id != RecordType.EOF); 
            //}
            //catch (Exception ex)
            //{
            //    TraceLogger.Error(ex.Message);
            //    TraceLogger.Debug(ex.ToString());
            //}
        }
    }
}
