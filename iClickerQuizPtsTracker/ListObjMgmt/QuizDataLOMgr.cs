using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using static iClickerQuizPtsTracker.AppConfigVals;

namespace iClickerQuizPtsTracker.ListObjMgmt
{
    public static class QuizDataLOMgr
    {
        #region fields
        private static int _noLOCols;
        #endregion

        #region pptys
        #region private
        private static QuizDataLOWrapper LOWrapperQuizData { get; set; }

        private static bool XLHdrRangesSet { get; set; }

        /// <summary>
        /// Gets a value indicating whether the <see cref="Excel.ListObject"/> of 
        /// quiz results yet contains any columns of quiz scores.
        /// </summary>
        /// <remarks>
        /// When a user first opens this workbook there will be no such columns.
        /// </remarks>
        private static bool HasDataCols { get; set; } = false;

        /// <summary>
        /// Gets the single-row <see cref="Excel.Range"/> of cells containing Session numbers.
        /// </summary>
        /// <remarks>
        /// This range spans the columns containing quiz scores.  As such, if 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper.HasDataCols"/> 
        /// is <see langword="false"/> then this range will not yet have been created.
        /// </remarks>
        private static Excel.Range XLRngSessNos { get; set; }

        /// <summary>
        /// Gets the single-row <see cref="Excel.Range"/> of cells containing the 
        /// dates on which quizzes were given.
        /// </summary>
        /// <remarks>
        /// This range spans the columns containing quiz scores.  As such, if 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper.HasDataCols"/> 
        /// is <see langword="false"/> then this range will not yet have been created.
        /// </remarks>
        private static Excel.Range XLRngSessDates { get; set; }

        /// <summary>
        /// Gets the single-row <see cref="Excel.Range"/> of cells identifying 
        /// which session (e.g., 1st, 2nd, etc.) within a given course-week 
        /// the quizzes were given.
        /// </summary>
        /// <remarks>
        /// This range spans the columns containing quiz scores.  As such, if 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper.HasDataCols"/> 
        /// is <see langword="false"/> then this range will not yet have been created.
        /// </remarks>
        private static Excel.Range XLRngWhichSess { get; set; }

        /// <summary>
        /// Gets the single-row <see cref="Excel.Range"/> of cells identifying 
        /// to which course-week the session corresponds.
        /// </summary>
        /// <remarks>
        /// This range spans the columns containing quiz scores.  As such, if 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper.HasDataCols"/> 
        /// is <see langword="false"/> then this range will not yet have been created.
        /// </remarks>
        private static Excel.Range XLRngCourseWk { get; set; }

        /// <summary>
        /// Gets the single-row <see cref="Excel.Range"/> of cells identifying 
        /// the maximum number of points that could be earned during 
        /// the Session&#39;s quiz.
        /// </summary>
        /// <remarks>
        /// This range spans the columns containing quiz scores.  As such, if 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper.HasDataCols"/> 
        /// is <see langword="false"/> then this range will not yet have been created.
        /// </remarks>
        private static Excel.Range XLRngMaxPts { get; set; }
        #endregion

        #region public
        /// <summary>
        /// Gets a <see cref="iClickerQuizPtsTracker.SortableBindingList{Session}"/> 
        /// containing all of the Sessions imported into the 
        /// <see cref="Globals.WshQuizPts"/> worksheet.
        /// </summary>
        public static SortableBindingList<Session> SortableBLSessions { get; private set; }

        public static Excel.ListObject XLListObj
        {
            get { return LOWrapperQuizData.XLTable; }
        }

        public static Excel.ListColumn XLListColEmls
        {
            get { return LOWrapperQuizData.XLTable.ListColumns[ColNmbrEmails]; }
        }

        public static Excel.ListColumn XLListColLastNm
        {
            get { return LOWrapperQuizData.XLTable.ListColumns[ColNmbrLastNms]; }
        }

        public static Excel.ListColumn XLListColFirstNm
        {
            get { return LOWrapperQuizData.XLTable.ListColumns[ColNmbrFirstNms]; }
        }

        /// <summary>
        /// Gets the total number of columns in the <see cref="Excel.ListObject"/>.
        /// </summary>
        public static int NmbrCols
        {
            get { return _noLOCols; }
        }

        /// <summary>
        /// Gets the number (index) of the column containing student email addresses.
        /// </summary>
        public static int ColNmbrEmails { get; private set; }

        /// <summary>
        /// /// Gets the number (index) of the column containing student last names.
        /// </summary>
        public static int ColNmbrLastNms { get; private set; }

        /// <summary>
        /// /// Gets the number (index) of the column containing student first names.
        /// </summary>
        public static int ColNmbrFirstNms { get; private set; }

        public static bool LOHasData
        {
            get { return LOWrapperQuizData.ListObjectHasData; }
        }
#endregion
        #endregion

        #region ctor
        static QuizDataLOMgr()
        {
            LOWrapperQuizData = new QuizDataLOWrapper(ThisWbkWrapper.LOInfoQuizData);
            // Both methods have Line 1 traps for whether they have already run...
            SetHasDataColsPpty();
            SetXLHdrRowRangesPptys();
        }
        #endregion

#region methods
        /// <summary>
        /// Adds a column to the Quiz Points <see cref="Excel.ListObject"/> for 
        /// a Session&#39;s quiz points, which also populating the column header 
        /// plus the appropriate cells above that column header with information 
        /// such as the Maximum Points, Quiz Date, etc.
        /// </summary>
        /// <param name="s">The <see cref="iClickerQuizPtsTracker.Session"/> for 
        /// which quiz scores will be added.</param>
        /// <param name="colNmbr">The number of the empty column which is being created.</param>
        /// <remarks>
        /// <list type="bullet">
        /// <item><description> The number of the column will never be less than 
        /// <see cref="iClickerQuizPtsTracker.AppConfigVals.DataTblNmbrRowLblCols"/> 
        /// plus 1.
        /// </description></item>
        /// <item><description>Empty columns are created keeping Session numbers 
        /// in order.  Session 01 will always be the right-hand-most column, and 
        /// the more recent Sessions will be the left-hand-most data columns.
        /// </description></item>
        /// </list>
        /// </remarks>
        public static void AddEmptyDataColumnWithHeaderInfo(Session s, out int colNmbr)
        {
            if (QuizDataLOMgr.SortableBLSessions == null)
                SetSortableBindingListPpty();

            colNmbr = 0; // ...to be tested
            byte nmbrTblCols = (byte)QuizDataLOMgr.LOWrapperQuizData.XLTable.Range.Columns.Count;

            if (!QuizDataLOMgr.HasDataCols)
            {
                colNmbr = nmbrTblCols + 1;
            }
            else
            {
                byte nmbrDataCols = (byte)(nmbrTblCols - DataTblNmbrRowLblCols);
                for (int i = 1; i <= nmbrDataCols; i++)
                {
                    if (s > SortableBLSessions[i])
                        colNmbr = i;
                }
                // Now trap for Sess No being smaller than any which are 
                // currently in quiz data table...
                if (colNmbr == 0)
                    colNmbr = nmbrTblCols + 1;
            }

            // Now either insert a column or expand table because our new 
            // Session data will be appended as the right-most column...
            if (colNmbr == nmbrTblCols + 1)
            {
                LOWrapperQuizData.XLTable.Resize(
                    LOWrapperQuizData.XLTable.Range.Resize[ColumnSize: nmbrTblCols + 1]);
            }
            else
            {
                Excel.Range col =  LOWrapperQuizData.WshParent.Columns[ColumnIndex: colNmbr];
                col.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }

            // Populate header info...
            int rngCol = colNmbr - DataTblNmbrRowLblCols;
            Excel.Range hdrCell;

            hdrCell = XLRngSessNos.Cells[1, rngCol];
            hdrCell.Value = s.SessionNo;

            hdrCell = XLRngSessDates.Cells[1, rngCol];
            hdrCell.Value = s.QuizDate;

            hdrCell = XLRngCourseWk.Cells[1, rngCol];
            hdrCell.Value = s.CourseWeek;

            hdrCell = XLRngMaxPts.Cells[1, rngCol];
            hdrCell.Value = s.MaxPts;

            hdrCell = XLRngWhichSess.Cells[1, rngCol];
            hdrCell.Value = Session.GetOrdinalNameFromWhichSessEnum(s.WeeklySession);

            hdrCell = LOWrapperQuizData.XLTable.HeaderRowRange.Cells[1, rngCol];
            hdrCell.Value = s.ColHeaderText;

            // Set the number of table columns field...
            _noLOCols = LOWrapperQuizData.XLTable.Range.Columns.Count;
        }

        /// <summary>
        /// Adds new students (emails, last names, & first names) to bottom of quiz data 
        /// <see cref="Excel.ListObject"/>, then resizes/redefines that <see cref="Excel.ListObject"/>.
        /// </summary>
        /// <param name="newStudentsList">A <see cref="System.Collections.Generic.List{Student}"/> 
        /// of students to be added.</param>
        /// <param name="isVirginWbk"><see langword="true"/>if we are dealing with a brand-new 
        /// workbook, otherwise <see langword="false"/>.</param>
        public static void AddAnyNewStudents(List<Student> newStudentsList)
        {
            int nmbrExistingStudents = 0; // ...default for virgin wbk
            int nmbrNewStudents = newStudentsList.Count;
            // Create & populate a 2D array of student info...
            object[,] arrxlStudentsToAdd = new object[nmbrNewStudents, 3];
            for (int i = 1; i <= nmbrNewStudents; i++)
            {
                arrxlStudentsToAdd[i - 1, ColNmbrEmails - 1] = newStudentsList[i - 1].EmailAddr;
                arrxlStudentsToAdd[i - 1, ColNmbrLastNms - 1] = newStudentsList[i - 1].LastName;
                arrxlStudentsToAdd[i - 1, ColNmbrFirstNms - 1] = newStudentsList[i - 1].FirstName;
            }

            // Define range where the new student info will be "pasted" (multi-step)...
            // We are assuming here that email column is still left-most column, and that 
            // last name & firt name columns are, in either order, the next 2 columns...
            Excel.Range rngAddStudents =
                LOWrapperQuizData.XLTable.ListColumns[ColNmbrEmails].DataBodyRange;
            rngAddStudents = rngAddStudents.Resize[RowSize: nmbrNewStudents, ColumnSize: 3];
            if (!ThisWbkWrapper.IsVirginWbk)
            {
                // We have to "re-locate" range to just below existing data body range...
                nmbrExistingStudents = LOWrapperQuizData.XLTable.DataBodyRange.Rows.Count;
                rngAddStudents = rngAddStudents.Offset[RowOffset: nmbrExistingStudents];
            }

            // "Paste" the new student info into the target range...
            rngAddStudents.Value = arrxlStudentsToAdd;

            // Re-define ListObject to encompass the rows of new students...
            Excel.Range newLO = LOWrapperQuizData.XLTable.HeaderRowRange.Resize[
                RowSize: nmbrExistingStudents + nmbrNewStudents + 1];
            LOWrapperQuizData.XLTable.Resize(newLO);
        }

        /// <summary>
        /// Populates the 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper.SortableBLSessions"/>.
        /// </summary>
        private static void SetSortableBindingListPpty()
        {
            if (!HasDataCols)
                return;
            if (SortableBLSessions == null)
                SortableBLSessions = new SortableBindingList<Session>();
            byte nmbrTblCols = (byte)LOWrapperQuizData.XLTable.Range.Columns.Count;
            for (byte i = 1; i <= nmbrTblCols - DataTblNmbrRowLblCols; i++)
            {
                byte wk = byte.Parse(XLRngCourseWk[1, i].Value);
                byte mxPts = byte.Parse(XLRngMaxPts[1, i].Value);
                DateTime qzDt = XLRngSessDates[1, i].Value2;
                string sessNo = (string)XLRngSessNos[1, i].Value;
                WkSession sessEnum =
                    Session.GetWhichSessEnumFromOrdinal(XLRngWhichSess[1, i].Value);
                SortableBLSessions.Add(new Session(sessNo, qzDt, mxPts, wk, sessEnum));
            }
        }

        private static void SetHasDataColsPpty()
        {
            byte nmbrTblCols = (byte)LOWrapperQuizData.XLTable.Range.Columns.Count;
            if (nmbrTblCols > DataTblNmbrRowLblCols)
                HasDataCols = true;
        }

        private static void SetXLHdrRowRangesPptys()
        {
            if (XLHdrRangesSet || !HasDataCols)
                return;
            else
            {
                byte nmbrTblCols = (byte)LOWrapperQuizData.XLTable.Range.Columns.Count;
                byte nmbrDataCols = (byte)(nmbrTblCols - DataTblNmbrRowLblCols);
                int rownoSessNos = LOWrapperQuizData.WshParent.Range["rowSessionNmbr"].Row;

                // 1st define Session Nos range as a single cell...
                XLRngSessNos = LOWrapperQuizData.WshParent.Cells[rownoSessNos, DataTblNmbrRowLblCols + 1];
                XLRngSessNos = XLRngSessNos.Resize[1, nmbrDataCols];

                // Define remaining ranges as offsets from _rngSessNos...
                XLRngCourseWk =
                    XLRngSessNos.Offset[LOWrapperQuizData.WshParent.Range["rowCourseWk"].Row - rownoSessNos];
                XLRngSessDates =
                    XLRngSessNos.Offset[LOWrapperQuizData.WshParent.Range["rowSessionDt"].Row - rownoSessNos];
                XLRngWhichSess =
                    XLRngSessNos.Offset[LOWrapperQuizData.WshParent.Range["rowSessionEnum"].Row - rownoSessNos];
                XLRngMaxPts =
                    XLRngSessNos.Offset[LOWrapperQuizData.WshParent.Range["rowTtlPts"].Row - rownoSessNos];
                XLHdrRangesSet = true;
            }
        }

        /// <summary>
        /// Populates the fields underlying the properties which get the number (index) 
        /// of columns containing student emails, last names, and first names.
        /// </summary>
        private static void SetStudentInfoColNmbrPptys()
        {
            ColNmbrEmails = LOWrapperQuizData.XLTable.ListColumns["Student ID"].Index;
            ColNmbrLastNms = LOWrapperQuizData.XLTable.ListColumns["Last Name"].Index;
            ColNmbrFirstNms = LOWrapperQuizData.XLTable.ListColumns["First Name"].Index;
        }
        #endregion
    }
}
