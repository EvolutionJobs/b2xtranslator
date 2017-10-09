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


namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// Class is / was used to extract different binary file internal files 
    /// </summary>
    //public class StreamExtractor
    //{
    //    StructuredStorageReader storageReader; 

    //    // some const declarations 
    //    public const string SUMMARYINFORMATIONSEARCH = "SummaryInformation";
    //    public const string DOCSUMMARYINFORMATIONSEARCH = "DocumentSummaryInformation";
    //    public const string WORKBOOKSEARCH = "Workbook"; 

    //    // Some attribute declarations 
    //    public string fileName;         /// The filename to the file to extract 
    //    ArrayList streamDirList;        // Streamnamelist 
    //    ArrayList streamList;           // collection of every other stream 

    //    // some stream attributes 
    //    public VirtualStream SummaryStream;
    //    public VirtualStream DocumentSummaryStream;
    //    public VirtualStream WorkbookStream; 

    //    /// <summary>
    //    /// Ctor with a file string 
    //    /// </summary>
    //    /// <param name="filename">The file to open</param>
    //    public StreamExtractor(string filename) 
    //    {
    //        if (filename == null)
    //        {
    //            throw new ExtractorException(ExtractorException.NOFILEFOUNDEXCEPTION); 
    //        }
    //        /// some initial variable declarations           
    //        this.fileName = filename;
    //        this.streamDirList = new ArrayList();
    //        this.streamList = new ArrayList();
    //        this.DocumentSummaryStream = null; 
    //        DateTime begin = DateTime.Now;
    //        storageReader = null; 
    //        try
    //        {
    //            // init StorageReader
    //            storageReader = new StructuredStorageReader(this.fileName);
    //            // read stream _entries
    //            ICollection<DirectoryEntry> streamEntries = storageReader.AllStreamEntries;

    //            foreach (object var in streamEntries)
    //            {
                    
    //                VirtualStream stream = storageReader.GetStream(((DirectoryEntry)var).Name);
    //                // checks which stream is read 
    //                if (((DirectoryEntry)var).Name.Contains(StreamExtractor.DOCSUMMARYINFORMATIONSEARCH)) 
    //                {
    //                    this.DocumentSummaryStream = stream; 
    //                }   
    //                else if (((DirectoryEntry)var).Name.Contains(StreamExtractor.SUMMARYINFORMATIONSEARCH)) 
    //                {
    //                    this.SummaryStream = stream; 
    //                }
    //                else if (((DirectoryEntry)var).Name.Contains(StreamExtractor.WORKBOOKSEARCH))
    //                {
    //                    this.WorkbookStream = stream;
    //                }
    //                else
    //                {
    //                    this.streamDirList.Add(var);
    //                    this.streamList.Add(stream);
    //                }
    //            }

    //        }
    //        catch (Exception ex)
    //        {
    //            throw new ExtractorException(ex.Message); 
    //        }
    //        finally
    //        {

    //        }

    //    }

    //    /// <summary>
    //    /// Overloaded ToString() method 
    //    /// </summary>
    //    /// <returns>String about this class </returns>
    //    public override string ToString()
    //    {
    //        string returnValue = "Extracted file: " + this.fileName + "\n Values: \n"; 
    //        // List filenames 
    //        // name and size of some streams 
    //        returnValue += "Stream: " + StreamExtractor.SUMMARYINFORMATIONSEARCH + "  Size: " + this.SummaryStream.Length + "\n";
    //        returnValue += "Stream: " + StreamExtractor.DOCSUMMARYINFORMATIONSEARCH + "  Size: " + this.DocumentSummaryStream.Length + "\n";
    //        returnValue += "Stream: " + StreamExtractor.WORKBOOKSEARCH + "  Size: " + this.WorkbookStream.Length + "\n"; 
    //        // all other not known streams 
    //        foreach (object var in streamDirList)
    //        {
    //            returnValue += ((DirectoryEntry)var).Name + "\n"; 
    //        }
    //        // List Streams 
    //        return returnValue;
    //    }

    //    /// <summary>
    //    /// Method is used to close the storage reader!
    //    /// </summary>
    //    public void closeFile()
    //    {

    //        if (storageReader != null)
    //        {
    //            storageReader.Close();
    //        }
    //    }

    //}
}
