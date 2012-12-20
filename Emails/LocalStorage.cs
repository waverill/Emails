using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using System.Windows.Forms;

namespace Emails
{
    /// <summary>
    /// This class facilitates local file storage for IAV e-mail project.
    /// </summary>
    /// <example>
    /// Command Line Application:
    /// <code>
    /// static void Main(string[] args)
    /// {
    ///     LocalStorage localStorage = new LocalStorage(".");
    ///     localStorage.ExtractZipFile(args[0], localStorage.CreateDateRangeName());
    ///     Console.WriteLine("Input File: {0}", args[0]);
    ///     Console.WriteLine("Extraction Directory: {0}", localStorage.ZipExtractDirectory);
    ///     if (localStorage.ZipFileList.Count > 0)
    ///         foreach (string zipFile in localStorage.ZipFileList)
    ///             Console.WriteLine(" - {0}", zipFile);
    ///     
    ///     Console.ReadLine();
    /// }
    /// </code>
    /// 
    /// Example Output:
    /// <code>
    /// C:\>application.exe "IAVM Releases 18 Aug 2011.zip"
    /// Input File: IAVM Releases 18 Aug 2011.zip
    /// Extraction Directory: .\20110822_20110826
    /// - 2011-b-0107_Symantec_EP.htm
    /// - 2011-a-0118_Mozilla.htm
    /// - 2011-a-0119_RealPlayer.htm
    /// - 2011-b-0105_ApacheTomcat.htm
    /// - 2011-b-0106_ISC_DHCP.htm
    /// - 2011-b-0108_Symantec_Veritas.htm
    /// </code>
    /// </example>
    class LocalStorage
    {
        // Class variables
        private List<string> _ZipFileList;
        private string _ZipDirectory;
        private string directory;

        /// <summary>
        /// Initializes a new instance of the LocalStorage class.
        /// </summary>
        /// <param name="baseDirectory">Directory to be used for storing data; directory must exist.</param>
        /// <exception cref="DirectoryNotFoundException"> if <paramref name="baseDirectory"/> does not exist</exception>
        public LocalStorage(string baseDirectory)
        {
            if (Directory.Exists(baseDirectory))
            {
                this.directory = baseDirectory;
                this._ZipFileList = new List<string>();
                this._ZipDirectory = "";
            }
            else
            {
                throw new DirectoryNotFoundException();
            }
        }

        /// <summary>
        /// Generates a date range string for the current week.
        /// </summary>
        /// <returns>A string containing the week's Monday date and Friday date ("YYYYMMDD_YYYYMMDD").</returns>
        /// <example>If the current date is 8/22/2011, then the return value would be "20110822_20110826".</example>
        /// <remarks>This will treat as <c>DayOfWeek.Monday</c> through <c>DayOfWeek.Friday</c> if the current date is <c>DayOfWeek.Saturday</c>.</remarks>
        public string CreateDateRangeName()
        {
            //DateTime dt_end = new DateTime(2011, 8, 27); // Used for testing
            DateTime dt_end = DateTime.Today;
            DayOfWeek curDayOfWeek = dt_end.DayOfWeek;

            // Handle fringe case should development be delayed until a Saturday
            if (dt_end.DayOfWeek == DayOfWeek.Saturday)
                dt_end = dt_end.Subtract(new TimeSpan(2, 0, 0, 0));

            while (curDayOfWeek != DayOfWeek.Friday)
            {
                dt_end = dt_end.AddDays(1);
                curDayOfWeek = dt_end.DayOfWeek;
            }
            DateTime dt_start = dt_end.Subtract(new TimeSpan(4, 0, 0, 0));

            string start_date = dt_start.Year.ToString() + dt_start.Month.ToString().PadLeft(2, '0') + dt_start.Day.ToString().PadLeft(2, '0');
            string end_date = dt_end.Year.ToString() + dt_end.Month.ToString().PadLeft(2, '0') + dt_end.Day.ToString().PadLeft(2, '0');

            return start_date + "_" + end_date;
        }

        /// <summary>
        /// Extracts the specified zip file to a subdirectory within <see cref="LocalStorage.baseDirectory"/>.
        /// </summary>
        /// <param name="file">The target zip file to extract.</param>
        /// <param name="subDirectory">A subdirectory within <see cref="LocalStorage.baseDirectory"/>; created if it doesn't exist.</param>
        /// <remarks>This will run basic sanity checks using private method <see cref="LocalStorage._ZipSanityChecks"/>.</remarks>
        public void ExtractZipFile(string file, string subDirectory)
        {
            _ZipSanityChecks(file);
            FastZip zipFile = new FastZip();
            zipFile.CreateEmptyDirectories = true;
            zipFile.RestoreDateTimeOnExtract = true;

            this._ZipDirectory = this.directory + "\\" + subDirectory;
            zipFile.ExtractZip(file, this._ZipDirectory, "");
        }

        /// <summary>
        /// Performs basic sanity checks to ensure file is intact and builds string list <see cref="LocalStorage._ZipFileList"/>.
        /// </summary>
        /// <param name="file">The zip file to examine.</param>
        /// <exception cref="FileNotFoundException"> if file does not exist.</exception>
        /// <exception cref="ZipException"> if file does not pass archive tests.</exception>
        /// <exception cref="ZipException"> if file contains no entries.</exception>
        /// <remarks>If all basic sanity checks pass then it builds string list <see cref="LocalStorage._ZipFileList"/>.</remarks>
        private void _ZipSanityChecks(string file)
        {
            MessageBox.Show(file);
            if (!File.Exists(file))
                throw new FileNotFoundException();

            using (ZipFile zf = new ZipFile(file))
            {
                // Basic sanity checks
                if (!zf.TestArchive(true))
                    throw new ZipException(file + " did not pass tests; possibly corrupted.");
                if (zf.Count == 0)
                    throw new ZipException(file + " contains no entries");

                // Store file list
                if (this._ZipFileList.Count == 0)
                    foreach (ZipEntry zEntry in zf)
                        this._ZipFileList.Add(zEntry.Name.ToString());
            }
        }

        /// <summary>
        /// Gets the files contained within a zip file.
        /// </summary>
        /// <remarks>Only returns the list if files have been extracted using <see cref="LocalStorage.ExtractZipFile"/>.</remarks>
        public List<string> ZipFileList
        {
            get { return _ZipFileList; }
        }

        /// <summary>
        /// Gets the extraction directory last used by <see cref="LocalStorage.ExtractZipFile"/>.
        /// </summary>
        public string ZipExtractDirectory
        {
            get { return this._ZipDirectory; }
        }
    }
}
