using b2xtranslator.DocFileFormat;
using b2xtranslator.StructuredStorage.Reader;
using NUnit.Framework;
using System;

namespace UnitTests
{
    [TestFixture]
    public class DocFileFormat
    {
        string file = @"files\simple.doc";
        StructuredStorageReader reader;
        WordDocument doc;



        [OneTimeSetUp]
        public void SetUp()
        {
            this.reader = new StructuredStorageReader(this.file);
            this.doc = new WordDocument(this.reader);
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            this.reader?.Close();
        }

        [Test]
        public void FirstCHPTest()
        {
            var chpx = this.doc.AllChpxFkps[0].grpchpx[0];
            var papx = this.doc.AllPapxFkps[0].grppapx[0];
            var chp = new CharacterProperties(chpx, papx, this.doc);
        }

        [Test]
        public void PieceTableTest()
        {
            Console.WriteLine("There are " + this.doc.PieceTable.Pieces.Count + " pieces in the table");

            foreach (var pcd in this.doc.PieceTable.Pieces)
            {
                //Console.WriteLine("\t"+pcd.cpStart + " - " + pcd.cpEnd + " : " + pcd.encoding.ToString() + " , starts at 0x" + String.Format("{0:x04}", pcd.fc));
                Console.WriteLine("Piece starts at "+ string.Format("{0:X04}", pcd.fc) + " and hast encoding "+pcd.encoding.ToString());
            }
        }

        [Test]
        public void DOPTest()
        {
            var dopBytes = new byte[(int)this.doc.FIB.lcbDop];
            this.doc.TableStream.Read(dopBytes, dopBytes.Length, (int)this.doc.FIB.fcDop);
            var dop = new DocumentProperties(this.doc.FIB, this.doc.TableStream);

            Console.WriteLine("Initial Footnote number: " + dop.nFtn);
        }

        /// <summary>
        /// prints the contents of the stylesheet
        /// </summary>
        [Test]
        public void STSHTest()
        {
            var stsh = new StyleSheet(this.doc.FIB, this.doc.TableStream, this.doc.DataStream);
            Console.WriteLine("Stylesheet contains " + stsh.Styles.Count + " Styles");

            for (int i=0; i<stsh.Styles.Count; i++)
            {
                Console.WriteLine("Style " + i);
                var std = stsh.Styles[i];
                if (std != null)
                {
                    Console.WriteLine("\tIdentifier: " + std.sti);
                    Console.WriteLine("\tStyle Kind: " + std.stk);
                    Console.WriteLine("\tBased On: " + std.istdBase); 
                    Console.WriteLine("\tName: " + std.xstzName);

                    if (std.papx != null)
                    {
                        Console.WriteLine("\t\tPAPX modifier:");
                        foreach (var sprm in std.papx.grpprl)
                        {
                            Console.WriteLine(string.Format("\t\tSPRM: modifies " + sprm.Type + " property 0x{0:x4} (" + sprm.Arguments.Length + " bytes)", (int)sprm.OpCode));
                        }
                    }

                    if (std.chpx != null)
                    {
                        Console.WriteLine("\t\tCHPX modifier:");
                        foreach (var sprm in std.chpx.grpprl)
                        {
                            Console.WriteLine(string.Format("\t\tSPRM: modifies " + sprm.Type + " property 0x{0:x4} (" + sprm.Arguments.Length + " bytes)", (int)sprm.OpCode));
                        }
                    }

                }
                else
                {
                    Console.WriteLine("\tEmpty Slot");
                }
            }
        }

        /// <summary>
        /// Method for testing FKP PAPX
        /// </summary>
        [Test]
        public void FKPPAPXTest()
        {
            //Get all PAPX FKPs
            var papxFkps = FormattedDiskPagePAPX.GetAllPAPXFKPs(this.doc.FIB, this.doc.WordDocumentStream, this.doc.TableStream, this.doc.DataStream);
            Console.WriteLine("There are " + papxFkps.Count + " FKPs with PAPX in this file: \n");
            foreach (var fkp in papxFkps)
            {
                Console.Write("FKP matches on " + fkp.crun + " paragraphs: ");
                foreach (int mark in fkp.rgfc)
                {
                    Console.Write(mark + " ");
                }
                Console.WriteLine("");
                for (int i = 0; i < fkp.crun; i++)
                {
                    var bx = fkp.rgbx[i];
                    var papx = fkp.grppapx[i];
                    Console.WriteLine("PAPX: has style " + papx.istd);
                    //foreach (SinglePropertyModifier sprm in papx.grpprl)
                    //{
                    //    Console.WriteLine(String.Format("\tSPRM: modifies " + sprm.Type + " property 0x{0:x4} (" + sprm.Arguments.Length + " bytes)", sprm.OpCode));
                    //}
                }
                Console.WriteLine("");
            }
        }

        /// <summary>
        /// Method for testing FKP CHPX
        /// </summary>
        [Test]
        public void FKPCHPXTest()
        {
            var chpxFkps = FormattedDiskPageCHPX.GetAllCHPXFKPs(this.doc.FIB, this.doc.WordDocumentStream, this.doc.TableStream);
            Console.WriteLine("There are " + chpxFkps.Count + " FKPs with CHPX in this file: \n");
            foreach (var fkp in chpxFkps)
            {
                Console.Write("FKP matches on " + fkp.crun + " characters: ");
                foreach (int mark in fkp.rgfc)
                {
                    Console.Write(mark + " ");
                }
                Console.WriteLine("");
                for (int i = 0; i < fkp.crun; i++)
                {
                    Console.WriteLine("CHPX:");
                    var chpx = fkp.grpchpx[i];
                    //foreach (SinglePropertyModifier sprm in chpx.grpprl)
                    //{
                    //    Console.WriteLine(String.Format("\tSPRM: modifies " + sprm.Type + " property 0x{0:x4} (" + sprm.Arguments.Length + " bytes)", sprm.OpCode));
                    //}
                }
            }
        }
    }
}
