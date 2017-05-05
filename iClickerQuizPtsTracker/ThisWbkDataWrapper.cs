using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using iClickerQuizPtsTracker.AppExceptions;
using iClickerQuizPtsTracker.Comparers;
using iClickerQuizPtsTracker.ListObjMgmt;
using static iClickerQuizPtsTracker.AppConfigVals;
using static iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper;

namespace iClickerQuizPtsTracker
{
    /// <summary>
    /// Provides a wrapper class for interacting with the <see cref="System.Data.DataTable"/> 
    /// of student quiz scores stored in this workbook.
    /// </summary>
    public static class ThisWbkDataWrapper
    {
        #region fields
        static Excel.Range _rngHdrzQzGradeCols;
        static DataTable _dtSessNos;
        static DataTable _dtEmls;
        static List<Student> _student = new List<Student>();
        static string _mostRecentQuizDt = "No data yet.";
        static string _mostRecentSessNos = " -- ";

        static BindingList<Session> _blSessions;

        #endregion

        #region ppts
        /// <summary>
        /// Gets a <see cref="System.Data.DataTable"/> of Session 
        /// Numbers that have been loaded into this workbook.
        /// </summary>
        public static DataTable SessionNmbrs
        {
            get { return _dtSessNos; }
        }

        /// <summary>
        /// Gets a <see cref="System.Data.DataTable"/> of email 
        /// addresses of all the students who have any quiz grade 
        /// activity at all loaded into this workbook.
        /// </summary>
        public static DataTable StudentEmails
        {
            get { return _dtEmls; }
        }

        public static List<Student> Students
        {
            get { return _student; }
        }

        /// <summary>
        /// Gets the most recent quiz date within the already-imported 
        /// quiz data.
        /// </summary>
        public static string MostRecentQuizDate
        {
            get  { return _mostRecentQuizDt; }
        }

        /// <summary>
        /// Gets the number of the most recent Session.
        /// </summary>
        /// <remarks>
        /// If there was more than one Session conducted on the most 
        /// recent date gets the numbers of all such Sessions.</remarks>
        public static string MostRecentSessNos
        {
            get  { return _mostRecentSessNos; }
        }

        /// <summary>
        /// Gets all <see cref="iClickerQuizPtsTracker.Session"/> objects
        /// in the <code>WshQuizPts</code> <see cref="Excel.Worksheet"/>.
        /// </summary>
        public static BindingList<Session> BListSession
        {
            get { return _blSessions; }
        }
        #endregion

        #region ctor
        /// <summary>
        /// Creates an instance of the <see cref="iClickerQuizPtsTracker.ThisWbkDataWrapper"/> class.
        /// </summary>
        static ThisWbkDataWrapper()
        {
            if(_loQzGrades.Range.Columns.Count > DataTblNmbrRowLblCols)
            {
                DefineQuizGradeColHeadersRange();
                PopulateSessionsBindingList();
                SetMostRecentSessDateNmbrsPptys();
            }
            
        }
        #endregion

        #region methods
        private static void DefineQuizGradeColHeadersRange()
        {
            int ttlTblCols = _loQzGrades.ListColumns.Count;
            // 2-step definition...
            _rngHdrzQzGradeCols = 
                _loQzGrades.HeaderRowRange.Resize[1, ttlTblCols - DataTblNmbrRowLblCols];
            _rngHdrzQzGradeCols = _rngHdrzQzGradeCols.Offset[0, DataTblNmbrRowLblCols];
        }
        
        private static void SetMostRecentSessDateNmbrsPptys()
        {
            SessionDateComparer dtComprr = new SessionDateComparer();
            Session newestSess = null;
            int comprsn;
            foreach(Session s in _blSessions)
            {
                comprsn = dtComprr.Compare(s, newestSess);
                // Note:  1st session will always be greater than null...
                switch (comprsn)
                {
                    case 1:
                        newestSess = s;
                        _mostRecentSessNos = s.SessionNo;
                        break;
                    case 0:
                        // We only need to add Session number...
                        _mostRecentSessNos =
                            string.Format($"{_mostRecentSessNos}, {s.SessionNo}");
                        break;
                    default:
                        break;
                }
            }
            _mostRecentQuizDt = newestSess.QuizDate.ToShortDateString();
        }
        
        /// <summary>
        /// Retreives all student emails from the iCLICKERQuizPoints worksheet.
        /// </summary>
        /// <returns>All student email in the &quot;Student ID&quot; column.</returns>
        public static IEnumerable<string> RetrieveStudentEmails()
        {
            Array arEmls = (Array)QuizDataLOMgr.XLListObj.ListColumns[QuizDataLOMgr.ColNmbrEmails].DataBodyRange;
            IEnumerable<string> _enumEmls = from string e in arEmls
                        orderby e
                        select e;
            return _enumEmls;
        }

        /// <summary>
        /// Retrieves the Session Numbers from the iCLICKERQuizPoints worksheet.
        /// </summary>
        /// <returns>
        /// All Session Numbers for which the worksheet has quiz scores.
        /// </returns>
        public static IEnumerable<string> RetrieveSessionNumbers()
        {
            Array arColHdrs = (Array)_loQzGrades.HeaderRowRange;
            IEnumerable<string> _enumSessionNos = from string h in arColHdrs
                                                  where (h.Contains("Session"))
                                                  orderby h
                                                  select h;
            return _enumSessionNos;
        }

        private static void PopulateSessionsBindingList()
        {
            int rowWeek = Globals.WshQuizPts.Range["rowCourseWk"].Row;
            int rowSessEnum = Globals.WshQuizPts.Range["rowSessionEnum"].Row;
            int rowSessNo = Globals.WshQuizPts.Range["rowSessionNmbr"].Row;
            int rowDt = Globals.WshQuizPts.Range["rowSessionDt"].Row;
            int rowMsxPts = Globals.WshQuizPts.Range["rowTtlQuizPts"].Row;

            foreach (Excel.Range c in _rngHdrzQzGradeCols)
            {
                int colNo = c.Column;
                string sessNo = $"{Globals.WshQuizPts.Cells[rowSessNo,colNo]}";

                Excel.Range cellDt = Globals.WshQuizPts.Cells[rowDt, colNo];
                DateTime dt = DateTime.Parse(cellDt.Value);

                Excel.Range cellPts = Globals.WshQuizPts.Cells[rowMsxPts, colNo];
                byte maxPts = byte.Parse(cellPts.Value);

                Excel.Range cellWkEnum = Globals.WshQuizPts.Cells[rowSessEnum, colNo];
                WkSession whichSess = Session.GetWhichSessEnumFromOrdinal(cellWkEnum.Value);

                byte courseWk = Globals.WshQuizPts.Cells[rowWeek, colNo];

                Session s = new Session(sessNo, dt, maxPts, courseWk, whichSess);
                _blSessions.Add(s);
            }
        }

        /// <summary>
        /// Creates and populates a <see cref="System.Data.DataTable"/> 
        /// which contains all of the Session information for quiz scores
        /// already imported into this workbook.
        /// </summary>
        public static void CreateSessionNosDataTable()
        {
            int rowOffset = Globals.WshQuizPts.Range["rowSessionNmbr"].Row -
                _loQzGrades.HeaderRowRange.Row;

            _dtSessNos = new DataTable("ThisWbkSessionNos");
            DataColumn colSessNo = new DataColumn("SessionNo", typeof(string));
            colSessNo.AllowDBNull = false;
            colSessNo.Unique = true;
            _dtSessNos.Columns.Add(colSessNo);

            foreach(Excel.Range c in _loQzGrades.HeaderRowRange)
            {
                if(string.Format($"{c.Value}").Contains("Session"))
                {
                    DataRow r = _dtSessNos.NewRow();
                    string sNo = ((string)c.Offset[rowOffset].Value).Trim();
                    if (sNo.Length == 1)
                        sNo = "0" + sNo; // ... pad with leading zero
                    r["SessionNo"] = sNo;
                    _dtSessNos.Rows.Add(r);
                }
            }
        }

        /// <summary>
        /// Creates and populates a <see cref="System.Data.DataTable"/> 
        /// which contains all student emails already imported 
        /// into this workbook.
        /// </summary>
        public static void CreateStudentEmailDataTable()
        {
            _dtEmls = new DataTable("ThisWbkEmails");
            DataColumn colEml = new DataColumn("StudentEml", typeof(string));
            colEml.AllowDBNull = false;
            colEml.Unique = true;
            _dtEmls.Columns.Add(colEml);

            Excel.Range em = 
                _loQzGrades.ListColumns[ColNmbrEmails].DataBodyRange;
            foreach(Excel.Range c in em)
            {
                DataRow r = _dtEmls.NewRow();
                r["StudentEml"] = (string)c.Value;
                _dtEmls.Rows.Add(r);
            }
        }

        /// <summary>
        /// Populates the <see cref="System.Collections.Generic.List{Student}"/> 
        /// field, which is exposed by the 
        /// <see cref="iClickerQuizPtsTracker.ThisWbkDataWrapper.Students"/> 
        /// property.
        /// </summary>
        public static void PopulateStudentList()
        {
            // Capture DataBodyRange values in an array...
            int nmbrRecs = _loQzGrades.DataBodyRange.Rows.Count;
            int nmbrCols = _loQzGrades.ListColumns.Count;
            object[,] arrxlDbr = new object[nmbrRecs, nmbrCols];
            arrxlDbr = _loQzGrades.DataBodyRange.Value;

            // Populate our student list...
            string eml;
            string lnm;
            string fnm;
            for(int i = 1; i<=_loQzGrades.DataBodyRange.Rows.Count; i++)
            {
                eml = arrxlDbr[i, ColNmbrEmails].ToString();
                lnm = arrxlDbr[i, ColNmbrLastNms].ToString();
                fnm = arrxlDbr[i, ColNmbrFirstNms].ToString();
                _student.Add(new Student(eml, lnm, fnm));
            }

        }

        /// <summary>
        /// Extracts the quiz date from a column header in the QuizPts worksheet.
        /// </summary>
        /// <param name="colHdr">A column header from the QuizPts worksheet.</param>
        /// <returns>The date of the quiz for the scores contained in the column.</returns>
        public static DateTime GetDatePortionFromColHeader(string colHdr)
        {
            int posHypen = colHdr.IndexOf("-");
            return DateTime.Parse(colHdr.Substring(posHypen + 1).Trim());
        }
        #endregion
    }
}
