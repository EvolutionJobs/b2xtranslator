using System;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.ZipUtils
{
	/// <summary>
	/// ZipReader defines an abstract class to read entries from a ZIP file.
	/// </summary>
	public abstract class ZipReader : IDisposable
	{
	    /// <summary>
	    /// Get an entry from a ZIP file.
	    /// </summary>
	    /// <param name="relativePath">The relative path of the entry in the ZIP
	    /// file.</param>
	    /// <returns>A stream containing the uncompressed data.</returns>
        public abstract Stream GetEntry(string relativePath);

        /// <summary>
        /// Close the ZIP file.
        /// </summary>
        public abstract void Close();

        #region IDisposable Members

        public virtual void Dispose()
        {
            this.Close();
        }

        #endregion
    }
}
