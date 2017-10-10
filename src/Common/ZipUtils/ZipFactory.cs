namespace b2xtranslator.ZipUtils
{
    /// <summary>
    /// ZipFactory provides instances of IZipReader and IZipWriter.
    /// </summary>
    public class ZipFactory
	{
	    /// <summary>
	    /// Provides an instance of IZipWriter.
	    /// </summary>
	    /// <param name="path">The path of the ZIP file to create.</param>
	    /// <returns></returns>
		public static ZipWriter CreateArchive(string path)
		{
			return new ZlibZipWriter(path);
		}
		
		/// <summary>
		/// Provides an instance of IZipReader.
		/// </summary>
		/// <param name="path">The path of the ZIP file to read.</param>
		/// <returns></returns>
		public static ZipReader OpenArchive(string path)
		{
			return new ZlibZipReader(path);
		}
	}
}
