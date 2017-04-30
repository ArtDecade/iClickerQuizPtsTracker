using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using iClickerQuizPtsTracker.AppExceptions;

namespace iClickerQuizPtsTracker
{
    /// <summary>
    /// Provides a centralized (static) method for accessing values in 
    /// the app.config file.
    /// </summary>
    public static class AppConfigVals
    {
        #region fields
        #region private
        private static List<String> _badKeys = new List<String>();
        #endregion

        #region public
        /// <summary>
        /// Gets the column number within raw iClicker files 
        /// containing student emails.
        /// </summary>
        public static readonly byte ExtFileColNoStudentEmail;

        /// <summary>
        /// Gets the column number within raw iClicker files 
        /// containing student names.
        /// </summary>
        public static readonly byte ExtFileColNoStudentName;

        /// <summary>
        /// Gets the number of columns of row labels within 
        /// raw iClicker files - i.e., the number of columns 
        /// <em>not</em> containing quiz scores.
        /// </summary>
        public static readonly byte ExtFileNmbrRowLblCols;

        /// <summary>
        /// Gets the number of columns of row labels within 
        /// the WshQuizPts sheet of this workboook - i.e., the 
        /// number of columns <em>not</em> containing quiz scores.
        /// </summary>
        public static readonly byte DataTblNmbrRowLblCols;

        /// <summary>
        /// Gets the value to be used for the
        /// <see cref="System.Data.DataColumn.ColumnName"/> 
        /// property of the <see cref="System.Data.DataColumn"/> 
        /// containing unique (primary key) IDs within the 
        /// <see cref="System.Data.DataTable"/> built from 
        /// a raw iClicker file.
        /// </summary>
        public static readonly string DataTblColNmID;

        /// <summary>
        /// Gets the value to be used for the
        /// <see cref="System.Data.DataColumn.ColumnName"/> 
        /// property of the <see cref="System.Data.DataColumn"/> 
        /// containing student emails within the 
        /// <see cref="System.Data.DataTable"/> built from 
        /// a raw iClicker file.
        /// </summary>
        public static readonly string DataTblColNmEmail;

        /// <summary>
        /// Gets the value to be used for the
        /// <see cref="System.Data.DataColumn.ColumnName"/> 
        /// property of the <see cref="System.Data.DataColumn"/> 
        /// containing student last names within the 
        /// <see cref="System.Data.DataTable"/> built from 
        /// a raw iClicker file.
        /// </summary>
        public static readonly string DataTblColNmLNm;

        /// <summary>
        /// Gets the value to be used for the
        /// <see cref="System.Data.DataColumn.ColumnName"/> 
        /// property of the <see cref="System.Data.DataColumn"/> 
        /// containing student first names within the 
        /// <see cref="System.Data.DataTable"/> built from 
        /// a raw iClicker file.
        /// </summary>
        public static readonly string DataTblColNmFNm;

        /// <summary>
        /// Gets the header/name of the <see cref="Excel.ListColumn"/> of student email 
        /// address within this workbook&#39;s Quiz Points <see cref="Excel.ListObject"/>.
        /// </summary>
        public static readonly string XLTblHdrEmail;

        /// <summary>
        /// Gets the header/name of the <see cref="Excel.ListColumn"/> of student last names 
        /// within this workbook&#39;s Quiz Points <see cref="Excel.ListObject"/>.
        /// </summary>
        public static readonly string XLTblHdrLName;

        /// <summary>
        /// Gets the header/name of the <see cref="Excel.ListColumn"/> of student first names 
        /// within this workbook&#39;s Quiz Points <see cref="Excel.ListObject"/>.
        /// </summary>
        public static readonly string XLTblHdrFName;

        /// <summary>
        /// Gets the header/name of the <see cref="Excel.ListColumn"/> of total quiz points 
        /// within this workbook&#39;s Quiz Points <see cref="Excel.ListObject"/>.
        /// </summary>
        public static readonly string XLTblHdrTtlPts;
        #endregion
        #endregion

        #region pptys
        /// <summary>
        /// Gets a <see cref="System.Collections.Generic.List{String}"/> 
        /// containing any keys within the <code>App.Config</code> 
        /// file which cannot be found by the 
        /// <see cref="System.Configuration.AppSettingsReader.GetValue(string, Type)"/> 
        /// method.
        /// </summary>
        public static List<String> BadAppConfigKeys
        {
            get
            {
                return _badKeys;
            }
        }
        #endregion

        #region ctor
        /// <summary>
        /// Initializes the <see cref="iClickerQuizPtsTracker.AppConfigVals"/> 
        /// static class.
        /// </summary>
        static AppConfigVals()
        {
            #region populateReadOnlyFields
            AppSettingsReader ar = new AppSettingsReader();

            // Populate fields from App.Config values.  If any exceptions add the 
            // corrupt key name to _badKeys...
            try
            {
                ExtFileColNoStudentEmail = (byte)ar.GetValue("ColNoEmailXL", typeof(byte));
            }
            catch
            {
                _badKeys.Add("ColNoEmailXL");
            }

            try
            {
                ExtFileColNoStudentName = (byte)ar.GetValue("ColNoStdntNmXL", typeof(byte));
            }
            catch
            {
                _badKeys.Add("ColNoStdntNmXL");
            }

            try
            {
                ExtFileNmbrRowLblCols = (byte)ar.GetValue("NmbrRowLblCols", typeof(Byte));
            }
            catch
            {
                _badKeys.Add("NmbrRowLblCols");
            }

            try
            {
                DataTblNmbrRowLblCols =
                    (byte)ar.GetValue("NmbrNonScoreCols", typeof(byte));
            }
            catch
            {
                _badKeys.Add("NmbrNonScoreCols");
            }

            try
            {
                DataTblColNmID = (string)ar.GetValue("ColID", typeof(string));
            }
            catch
            {
                _badKeys.Add("ColID");
            }

            try
            {
                DataTblColNmEmail = (string)ar.GetValue("ColEmail", typeof(string));
            }
            catch
            {
                _badKeys.Add("ColEmail");
            }

            try
            {
                DataTblColNmFNm = (string)ar.GetValue("ColFN", typeof(string));
            }
            catch
            {
                _badKeys.Add("ColFN");
            }

            try
            {
                DataTblColNmLNm = (string)ar.GetValue("ColLN", typeof(string));
            }
            catch
            {
                _badKeys.Add("ColLN");
            }

            try
            {
                XLTblHdrEmail = (string)ar.GetValue("QuizDataXLTblHdrEml", typeof(string));
            }
            catch
            {
                _badKeys.Add("QuizDataXLTblHdrEml");
            }

            try
            {
                XLTblHdrLName = (string)ar.GetValue("QuizDataXLTblHdrLNm", typeof(string));
            }
            catch
            {
                _badKeys.Add("QuizDataXLTblHdrLNm");
            }

            try
            {
                XLTblHdrFName = (string)ar.GetValue("QuizDataXLTblHdrFNm", typeof(string));
            }
            catch
            {
                _badKeys.Add("QuizDataXLTblHdrFNm");
            }

            try
            {
                XLTblHdrTtlPts = (string)ar.GetValue("QuizDataXLTblHdrTtlPts", typeof(string));
            }
            catch
            {
                _badKeys.Add("QuizDataXLTblHdrTtlPts");
            }
            #endregion

            // Now check for errors...
            if (_badKeys.Count > 0)
            {
                InalidAppConfigItemException ex = 
                    new InalidAppConfigItemException();
                throw ex;
            }
        }
        #endregion
    }
}
