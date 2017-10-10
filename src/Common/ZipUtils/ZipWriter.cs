using System.IO;

namespace b2xtranslator.ZipUtils
{
    /// <summary>
    /// ZipWriter defines an abstract class to write entries into a ZIP file.
    /// To add a file, first call AddEntry with the relative path, then
    /// write the content of the file into the stream.
    /// </summary>
    public abstract class ZipWriter : Stream
	{
        public string Filename = "";

	    /// <summary>
	    /// Adds an entry to the ZIP file (only writes the header, to write
	    /// the content use Stream methods).
	    /// </summary>
	    /// <param name="relativePath">The relative path of the entry in the ZIP
	    /// file.</param>
	    public abstract void AddEntry(string relativePath);
	}
}
