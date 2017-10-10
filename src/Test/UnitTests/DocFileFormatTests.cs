using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using System.IO;
using System.Xml;
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace UnitTests
{
    [TestFixture]
    public class DocFileFormatTests
    {
        private Application word2007 = null;
        private List<FileInfo> files;
        private Object confirmConversions = Type.Missing;
        private Object readOnly = true;
        private Object addToRecentFiles = Type.Missing;
        private Object passwordDocument = Type.Missing;
        private Object passwordTemplate = Type.Missing;
        private Object revert = Type.Missing;
        private Object writePasswordDocument = Type.Missing;
        private Object writePasswordTemplate = Type.Missing;
        private Object format = Type.Missing;
        private Object encoding = Type.Missing;
        private Object visible = Type.Missing;
        private Object openConflictDocument = Type.Missing;
        private Object openAndRepair = Type.Missing;
        private Object documentDirection = Type.Missing;
        private Object noEncodingDialog = Type.Missing;

        private Object saveChanges = false;
        private Object originalFormat = false;
        private Object routeDocument = false;


        [TestFixtureSetUp]
        public void SetUp()
        {
            //read the config
            var fs = new FileStream("Config.xml", FileMode.Open);
            var config = new XmlDocument();
            config.Load(fs);
            fs.Close();

            //read the inputfiles
            this.files = new List<FileInfo>();
            foreach (XmlNode fileNode in config.SelectNodes("input-files/file"))
            {
                this.files.Add(new FileInfo(fileNode.Attributes["path"].Value));
            }

            //start the application
            this.word2007 = new Application();
        }


        [TestFixtureTearDown]
        public void TearDown()
        {
            this.word2007.Quit(
                ref saveChanges,
                ref originalFormat,
                ref routeDocument);
        }


        /// <summary>
        /// Tests if the inputfile is parsable
        /// </summary>
        [Test]
        public void TestParseability()
        {
            foreach (var inputFile in this.files)
            {
                try
                {
                    var reader = new StructuredStorageReader(inputFile.FullName);
                    var doc = new WordDocument(reader);
                    Console.WriteLine("PASSED TestParseability " + inputFile.FullName);
                }
                catch (Exception e)
                {
                    throw new AssertionException(e.Message + inputFile.FullName, e);
                }
            }
        }


        [Test]
        public void TestProperties()
        {
            foreach (var inputFile in this.files)
            {
                var omDoc = loadDocument(inputFile.FullName);
                var dffDoc = new WordDocument(new StructuredStorageReader(inputFile.FullName));

                string dffRevisionNumber = dffDoc.DocumentProperties.nRevision.ToString();
                string omRevisionNumber = (string)getDocumentProperty(omDoc, "Revision number");

                object omCreationDate = (DateTime)getDocumentProperty(omDoc, "Creation date");
                var dffCreationDate = dffDoc.DocumentProperties.dttmCreated.ToDateTime();
                
                object omLastPrintedDate = getDocumentProperty(omDoc, "Last print date");
                var dffLastPrintedDate = dffDoc.DocumentProperties.dttmLastPrint.ToDateTime();

                omDoc.Close(ref saveChanges, ref originalFormat, ref routeDocument);

                try
                {
                    Assert.AreEqual(omRevisionNumber, dffRevisionNumber);

                    if (omCreationDate != null && ((DateTime)omCreationDate).Year != 1601)
                    {
                        Assert.AreEqual((DateTime)omCreationDate, dffCreationDate);
                    }

                    if (omLastPrintedDate != null && ((DateTime)omLastPrintedDate).Year != 1601)
                    {
                        Assert.AreEqual((DateTime)omLastPrintedDate, dffLastPrintedDate);
                    }

                    Console.WriteLine("PASSED TestProperties " + inputFile.FullName);
                }
                catch (AssertionException e)
                {
                    throw new AssertionException(e.Message + inputFile.FullName, e);
                }
            }
        }


        /// <summary>
        /// 
        /// </summary>
        public void TestCharacters()
        { 
            foreach (var inputFile in this.files)
            {
                var omDoc = loadDocument(inputFile.FullName);
                var dffDoc = new WordDocument(new StructuredStorageReader(inputFile.FullName));
                omDoc.Fields.ToggleShowCodes();

                var omText = new StringBuilder();
                var omMainText = omDoc.StoryRanges[WdStoryType.wdMainTextStory].Text.ToCharArray();
                foreach(char c in omMainText)
                {
                    if ((int)c > 0x20)
                    {
                        omText.Append(c);
                    }
                }

                var dffText = new StringBuilder();
                var dffMainText = dffDoc.Text.GetRange(0, dffDoc.FIB.ccpText);
                foreach (char c in dffMainText)
                {
                    if ((int)c > 0x20)
                    {
                        dffText.Append(c);
                    }
                }

                try
                {
                    Assert.AreEqual(omText.ToString(), dffText.ToString());

                    Console.WriteLine("PASSED TestCharacters " + inputFile.FullName);
                }
                catch (AssertionException e)
                {
                    throw new AssertionException(e.Message + inputFile.FullName, e);
                }
            }
        }


        /// <summary>
        /// Tests the count of bookmarks in the documents.
        /// Also tests the start and the end position a randomly selected bookmark.
        /// </summary>
        [Test]
        public void TestBookmarks()
        {
            foreach (var inputFile in this.files)
            {
                var omDoc = loadDocument(inputFile.FullName);
                var dffDoc = new WordDocument(new StructuredStorageReader(inputFile.FullName));
                omDoc.Bookmarks.ShowHidden = true;

                int omBookmarkCount = omDoc.Bookmarks.Count;
                int dffBookmarkCount = dffDoc.BookmarkNames.Strings.Count;
                int omBookmarkStart = 0;
                int dffBookmarkStart = 0;
                int omBookmarkEnd = 0;
                int dffBookmarkEnd = 0;

                if (omBookmarkCount > 0 && dffBookmarkCount > 0)
                {
                    //generate a randomly selected bookmark
                    var rand = new Random();
                    object omIndex = rand.Next(0, dffBookmarkCount);

                    //get the index's bookmark
                    var omBookmark = omDoc.Bookmarks.get_Item(ref omIndex);
                    omBookmarkStart = omBookmark.Start;
                    omBookmarkEnd = omBookmark.End;

                    //get the bookmark with the same name from DFF
                    int dffIndex = 0;
                    for (int i = 0; i < dffDoc.BookmarkNames.Strings.Count; i++)
                    {
                        if (dffDoc.BookmarkNames.Strings[i] == omBookmark.Name)
                        {
                            dffIndex = i;
                            break;
                        }
                    }
                    dffBookmarkStart = dffDoc.BookmarkStartPlex.CharacterPositions[dffIndex];
                    dffBookmarkEnd = dffDoc.BookmarkEndPlex.CharacterPositions[dffIndex];
                }

                omDoc.Close(ref saveChanges, ref originalFormat, ref routeDocument);

                try
                {
                    //compare bookmark count
                    Assert.AreEqual(omBookmarkCount, dffBookmarkCount);

                    //compare bookmark start
                    Assert.AreEqual(omBookmarkStart, dffBookmarkStart);

                    //compare bookmark end
                    Assert.AreEqual(omBookmarkEnd, dffBookmarkEnd);

                    Console.WriteLine("PASSED TestBookmarks " + inputFile.FullName);
                }
                catch (AssertionException e)
                {
                    throw new AssertionException(e.Message + inputFile.FullName, e);
                }
            }
        }
        

        /// <summary>
        /// Tests the count of of comments in the documents.
        /// Also compares the author of the first comment.
        /// </summary>
        [Test]
        public void TestComments()
        {
            foreach (var inputFile in this.files)
            {
                var omDoc = loadDocument(inputFile.FullName);
                var dffDoc = new WordDocument(new StructuredStorageReader(inputFile.FullName));

                int dffCommentCount = dffDoc.AnnotationsReferencePlex.Elements.Count;
                int omCommentCount = omDoc.Comments.Count;
                string omFirstCommentInitial = "";
                string omFirstCommentAuthor = "";
                string dffFirstCommentInitial = "";
                string dffFirstCommentAuthor = "";

                if (dffCommentCount > 0 && omCommentCount > 0)
                {
                    var omFirstComment = omDoc.Comments[1];
                    var dffFirstComment = (AnnotationReferenceDescriptor)dffDoc.AnnotationsReferencePlex.Elements[0];

                    omFirstCommentInitial = omFirstComment.Initial;
                    omFirstCommentAuthor = omFirstComment.Author;

                    dffFirstCommentInitial = dffFirstComment.UserInitials;
                    dffFirstCommentAuthor = dffDoc.AnnotationOwners[dffFirstComment.AuthorIndex];
                }
                
                omDoc.Close(ref saveChanges, ref originalFormat, ref routeDocument);

                try
                {
                    //compare comment count
                    Assert.AreEqual(omCommentCount, dffCommentCount);

                    //compare initials
                    Assert.AreEqual(omFirstCommentInitial, dffFirstCommentInitial);

                    //compate the author names
                    Assert.AreEqual(omFirstCommentAuthor, dffFirstCommentAuthor);

                    Console.WriteLine("PASSED TestComments " + inputFile.FullName);
                }
                catch (AssertionException e)
                {
                    throw new AssertionException(e.Message + inputFile.FullName, e);
                }
            }
        }

        private Document loadDocument(Object filename)
        {
            return this.word2007.Documents.Open(
                ref filename,
                ref confirmConversions,
                ref readOnly,
                ref addToRecentFiles,
                ref passwordDocument,
                ref passwordTemplate,
                ref revert,
                ref writePasswordDocument,
                ref writePasswordTemplate,
                ref format,
                ref encoding,
                ref visible,
                ref openConflictDocument,
                ref openAndRepair,
                ref documentDirection,
                ref noEncodingDialog);
        }

        private object getDocumentProperty(Document document, string propertyName)
        {
            object propertyValue = null;
            try
            {
                object builtInProperties = document.BuiltInDocumentProperties;
                var builtInPropertiesType = builtInProperties.GetType();

                object property = builtInPropertiesType.InvokeMember("Item", BindingFlags.GetProperty, null, builtInProperties, new object[] { propertyName });
                var propertyType = property.GetType();

                propertyValue = propertyType.InvokeMember("Value", BindingFlags.GetProperty, null, property, new object[] { });
            }
            catch (TargetInvocationException)
            {
            }

            return propertyValue;
        }

    }
}
