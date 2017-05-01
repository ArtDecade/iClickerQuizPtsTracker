using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.ComponentModel;
using iClickerQuizPtsTracker.ListObjMgmt;
using iClickerQuizPtsTracker.Comparers;
using static iClickerQuizPtsTracker.ThisWbkDataWrapper;
using static iClickerQuizPtsTracker.AppConfigVals;

namespace iClickerQuizPtsTracker
{
    /// <summary>
    /// Provides a mechanism for  with the workbook&apos;s action panel.
    /// </summary>
    public static class UserControlsHandler
    {
        #region fields
        private static byte _crsWk = 0;
        private static WkSession _session = WkSession.None;
        private static DataTable _dtSortedSsnsAll;
        private static EPPlusManager _eppMgr;
        private static BindingList<Session> _blAllSessns = new BindingList<Session>();
        private static BindingList<Session> _blNewSessns = new BindingList<Session>();
        private static bool _newStudentsAdded = false;
        #endregion

        #region Ppts
        /// <summary>
        /// Gets the <see cref="iClickerQuizPtsTracker.WkSession"/> enumeration indicating
        /// which session/recitation within a given semester week the to-be-imported column
        /// of student quiz scores represents.
        /// </summary>
        public static WkSession WhichSession
        {
            get
            { return _session; }
        }

        /// <summary>
        /// Gets the semester week in which the to-be-imported column of student
        /// quiz scores occurred.
        /// </summary>
        public static byte CourseWeek
        {
            get
            { return _crsWk; }
        }

        /// <summary>
        /// Gets or sets the date on which the to-be-imported column of student quiz
        /// scores occurred.
        /// </summary>
        public static DateTime QuizDate { get; set; } = DateTime.Parse("1/1/2016");

        /// <summary>
        /// Gets a <see cref="System.ComponentModel.BindingList{Session}"/> of 
        /// all sessions in the raw quiz file.
        /// </summary>
        public static BindingList<Session> BindingListAllSessions
        {
            get
            { return _blAllSessns; }
        }

        /// <summary>
        /// Gets a <see cref="System.ComponentModel.BindingList{Session}"/> of 
        /// data in the raw quiz file that have not yet been imported into 
        /// iCLICKERQuizPoints worksheet.
        /// </summary>
        public static BindingList<Session> BindingListNewSessions
        {
            get
            { return _blNewSessns; }
        }
        #endregion

        #region methods
        /// <summary>
        /// Sets the <see cref="iClickerQuizPtsTracker.UserControlsHandler.CourseWeek"/> property.
        /// </summary>
        /// <param name="selectedWk">The week of the semester in which the to-be-imported column of student quiz
        /// scores occurred.</param>
        public static void SetCourseWeek(string selectedWk)
        {
            _crsWk = byte.Parse(selectedWk);
        }

        /// <summary>
        /// Sets the <see cref="iClickerQuizPtsTracker.WkSession"/> property.
        /// </summary>
        /// <param name="session">Which session within a semester week represented by the 
        /// to-be-imported column of data.</param>
        public static void SetSessionEnum(string session)
        {
            switch (session)
            {
                case "First":
                    _session = WkSession.First;
                    break;
                case "Second":
                    _session = WkSession.Second;
                    break;
                case "Third":
                    _session = WkSession.Third;
                    break;
                default:
                    _session = WkSession.None;
                    break;
            }
        }

        /// <summary>
        /// Fires all other methods required to import data from an Excel file of 
        /// raw iClicker student test scores.
        /// </summary>
        public static void MaestroOpenDataFile()
        {
            string rawDataFileFullNm;
            bool userSelectedFile;
            userSelectedFile = PromptUserToOpenQuizDataWbk(out rawDataFileFullNm);
            if (!userSelectedFile)
                return;
            // If here user selected a file...
            _eppMgr = new EPPlusManager(rawDataFileFullNm);
            _eppMgr.CreateRawQuizDataTable();

            // Set BindingList of all Sessions...
            _blAllSessns = _eppMgr.BListSessionsAll;

            // Get BindingList of existing Sessions...
            BindingList<Session> blExisting = ThisWbkDataWrapper.BListSession;

            // Create BindingList of new Sessions...
            if (blExisting == null)
            {
                foreach (Session s in _blAllSessns)
                    _blNewSessns.Add(s);
            }
            else
            {
                var newSessns = (from Session sAll
                                 in _eppMgr.BListSessionsAll
                                 orderby sAll.SessionNo
                                 select sAll).Except(from Session sExisting
                                                     in blExisting
                                                     select sExisting);
                foreach (Session s in newSessns)
                    _blNewSessns.Add(s);
            }
        }

        public static void MaestroImportSessionData(string sessNo, byte crsWk, 
            WkSession whichSess)
        {
            // Use session number to find session in raw data file...
            Session sessToImport = null;
            foreach(Session s in _blAllSessns)
            {
                if(s.SessionNo == sessNo)
                {
                    sessToImport = s;
                    break;
                }
            }

            // Set properties passed in from user control panel...
            sessToImport.CourseWeek = crsWk;
            sessToImport.WeeklySession = whichSess;

            // Check that user is not skipping a session...
            if(!HasUserImportedPriorSessionsForWk(sessToImport))
            {
                MsgBoxGenerator.SetMissingPriorSessionWithinWkMsg();
                MsgBoxGenerator.ShowMsg(System.Windows.Forms.MessageBoxButtons.OK);
                return; // ... we gotta stop
            }

            // Do our magic...
            int colnoNewSess;
            QuizDataLOWrapper.AddEmptyDataColumnWithHeaderInfo(
                sessToImport, out colnoNewSess);
            if(!_newStudentsAdded)
                AddAnyNewStudentsToThisWbk();
            ImportQuizData(sessToImport, colnoNewSess);

            // Add a TA-DA message
        }

        /// <summary>
        /// Prompts user to select the Excel containing latest iClick data.
        /// </summary>
        /// <param name="dataFileFullNm">An <code>out</code> parameter to 
        /// capture the name of the selected file.</param>
        /// <returns>
        /// <see langword="true"/> if the user selected a file, otherwise
        /// <see langword="false"/>.
        /// </returns>
        /// <remarks>
        /// If the user does not select a file then the <code>dataFileFullNm</code> 
        /// out parameter will be set to <see cref="string.Empty"/>.
        /// </remarks>
        private static bool PromptUserToOpenQuizDataWbk(out string dataFileFullNm)
        {
            dataFileFullNm = string.Empty; // ...in case user cxls
            bool userSelectedWbk = new bool();

            Office.FileDialog fd = Globals.ThisWorkbook.Application.get_FileDialog(
                Office.MsoFileDialogType.msoFileDialogFilePicker);
            fd.Title = "Latest iClick Results";
            fd.AllowMultiSelect = false;
            fd.Filters.Clear();
            fd.Filters.Add("Excel Files", "*.xlsx");

            // Handle user selection...
            if (fd.Show() == -1) // ...-1 == file selected; 0 == user cxled
            {
                userSelectedWbk = true;
                dataFileFullNm = fd.SelectedItems.Item(1);
            }
            return userSelectedWbk;
        }


        private static bool HasUserImportedPriorSessionsForWk(Session sessToImport)
        {
            bool havePriorSessnsForWk = false;
            // Create a list of already-imported Sessions which come from the same 
            // course week as the Session we want to import...
            SessionCourseWkComparer cwComprr = new SessionCourseWkComparer();
            List<Session> thisWkSessions = new List<Session>();
            foreach(Session importedSess in ThisWbkDataWrapper.BListSession)
            {
                if (cwComprr.Compare(sessToImport, importedSess) == 0)
                    thisWkSessions.Add(importedSess);
            }

            switch(sessToImport.WeeklySession)
            {
                case WkSession.First:
                    havePriorSessnsForWk = true;
                    break;
                case WkSession.Second:
                    // if thisWkSessions.Count != 1 havePriorSessnsForWk remains false...
                    if (thisWkSessions.Count == 1)
                    {
                        if (ListContainsSessionWithWkSession(thisWkSessions, WkSession.First))
                            havePriorSessnsForWk = true;
                    }
                    break;
                case WkSession.Third:
                    // if thisWkSessions.Count != 2 havePriorSessnsForWk remains false...
                    if (thisWkSessions.Count == 2)
                    {
                        if (ListContainsSessionWithWkSession(thisWkSessions, WkSession.First) &&
                            ListContainsSessionWithWkSession(thisWkSessions, WkSession.Second))
                        { havePriorSessnsForWk = true; }
                    }
                    break;
            }
            return havePriorSessnsForWk;
        }

        private static bool ListContainsSessionWithWkSession(List<Session> listToCk, WkSession sessEnum)
        {
            bool hasSess = false;
            foreach(Session s in listToCk)
            {
                if(s.WeeklySession == sessEnum)
                {
                    hasSess = true;
                    break;
                }
            }
            return hasSess;
        }

        private static void AddAnyNewStudentsToThisWbk()
        {
            if (ThisWbkWrapper.IsVirginWbk)
            {
                QuizDataLOWrapper.AddAnyNewStudents(_eppMgr.Students);
            }
            else
            {
                PopulateStudentList();
                // Get all students in external data file but not in this wbk...
                var studentsToAdd =
                    ((from Student s in _eppMgr.Students select s)
                        .Except(from Student s in ThisWbkDataWrapper.Students
                                select s)).AsEnumerable();
                List<Student> newStudents = new List<Student>();
                foreach (Student st in studentsToAdd)
                    newStudents.Add(st);
                if (newStudents.Count > 0)
                    QuizDataLOWrapper.AddAnyNewStudents(newStudents);
            }
            PopulateStudentList(); // ...to redefine
            // Update flag!!!!...
            _newStudentsAdded = true;
        }

        private static void ImportQuizData(Session s, int colNoNewSession)
        {
            // TODO - Verify we update Students List after importing any new students...
            int i = 0;
            int nmbrStdnts = ThisWbkDataWrapper.Students.Count;
            string dataColNm = ReconstructRawDataColumnHeader(s);
            object[,] arrxlQzGrds = new object[nmbrStdnts, 1];
            foreach(Student st in ThisWbkDataWrapper.Students)
            {
                // Build our array of scores to import...
                byte? sc = (byte)((from r in _eppMgr.RawQuizScoresDataTable.AsEnumerable()
                         where r.Field<string>(DataTblColNmEmail) == st.EmailAddr 
                         select r).First())[dataColNm];
                arrxlQzGrds[i, 1] = sc;
                i++;
            }

            // "Paste" the array of data into our column...
            ThisWbkDataWrapper.XLTblQuizData.ListColumns[colNoNewSession].DataBodyRange.Value =
                arrxlQzGrds;

        }

        private static string ReconstructRawDataColumnHeader(Session s)
        {
            // Sample:  Session 4 Total 5/2/16 [2.00]...
            string sessNo = byte.Parse(s.SessionNo).ToString();
            string dt = s.QuizDate.ToString("m/d/yy");
            string maxPts = s.MaxPts.ToString("0.00");
            maxPts = string.Format($"[{maxPts}]");

            return string.Format($"Session {sessNo} Total {dt} {maxPts}");
        }

        #endregion
    }
}
