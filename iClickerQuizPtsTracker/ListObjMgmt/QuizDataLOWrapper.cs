﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using static iClickerQuizPtsTracker.AppConfigVals;

namespace iClickerQuizPtsTracker.ListObjMgmt
{
    /// <summary>
    /// Provides a mechanism for interacting with the workbook's <see cref="Excel.ListObject"/> 
    /// of iClicker quiz data.
    /// </summary>
    public class QuizDataLOWrapper : XLListObjWrapper
    {
        #region fields
        private static WshListobjPair _wshTblPair;
        private static int _noLOCols;
        private static int _colnoEmls;
        private static int _colnoFNms;
        private static int _colnoLNms;
        private static bool _staticPptsSet = false;
        private static bool _hasDataCols;
        private static Excel.Range _rngSessNos;
        private static Excel.Range _rngSessDts;
        private static Excel.Range _rngSessEnums;
        private static Excel.Range _rngCourseWk;
        private static Excel.Range _rngMaxPts;
        private static SortableBindingList<Session> _srtblBLSessns;
        
        #endregion

        #region ppts
        /// <summary>
        /// Gets a value indicating whether the <see cref="Excel.ListObject"/> of 
        /// quiz results yet contains any columns of quiz scores.
        /// </summary>
        /// <remarks>
        /// When a user first opens this workbook there will be no such columns.
        /// </remarks>
        public static bool HasDataCols
        {
            get { return _hasDataCols; }
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
        public static int ColNmbrEmails
        {
            get { return _colnoEmls; }
        }

        /// <summary>
        /// /// Gets the number (index) of the column containing student last names.
        /// </summary>
        public static int ColNmbrLastNms
        {
            get { return _colnoLNms; }
        }

        /// <summary>
        /// /// Gets the number (index) of the column containing student first names.
        /// </summary>
        public static int ColNmbrFirstNms
        {
            get { return _colnoFNms; }
        }

        /// <summary>
        /// Gets the single-row <see cref="Excel.Range"/> of cells containing Session numbers.
        /// </summary>
        /// <remarks>
        /// This range spans the columns containing quiz scores.  As such, if 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper.HasDataCols"/> 
        /// is <see langword="false"/> then this range will not yet have been created.
        /// </remarks>
        public static Excel.Range XLRngSessNos
        {
            get { return _rngSessNos; }
        }

        /// <summary>
        /// Gets the single-row <see cref="Excel.Range"/> of cells containing the 
        /// dates on which quizzes were given.
        /// </summary>
        /// <remarks>
        /// This range spans the columns containing quiz scores.  As such, if 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper.HasDataCols"/> 
        /// is <see langword="false"/> then this range will not yet have been created.
        /// </remarks>
        public static Excel.Range XLRngSessDates
        {
            get { return _rngSessDts; }
        }

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
        public static Excel.Range XLRngWhichSess
        {
            get { return _rngSessEnums; }
        }

        /// <summary>
        /// Gets the single-row <see cref="Excel.Range"/> of cells identifying 
        /// to which course-week the session corresponds.
        /// </summary>
        /// <remarks>
        /// This range spans the columns containing quiz scores.  As such, if 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper.HasDataCols"/> 
        /// is <see langword="false"/> then this range will not yet have been created.
        /// </remarks>
        public static Excel.Range XLRngCourseWk
        {
            get { return _rngCourseWk; }
        }

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
        public static Excel.Range XLRngMaxPts
        {
            get { return _rngMaxPts; }
        }

        /// <summary>
        /// Gets a <see cref="iClickerQuizPtsTracker.SortableBindingList{Session}"/> 
        /// containing all of the Sessions imported into the 
        /// <see cref="Globals.WshQuizPts"/> worksheet.
        /// </summary>
        public static SortableBindingList<Session> SortableBLSessions
        {
            get { return _srtblBLSessns; }
        }
        #endregion

        #region ctor
        /// <summary>
        /// Initializes a new instance of the 
        /// class <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper"/>.
        /// </summary>
        /// <param name="wshTblNmzPair">The properties of this <see langword="struct"/> 
        /// should be populated with the name of the <see cref="Excel.ListObject"/> 
        /// containing the iClicker quiz data and the name 
        /// of the parent <see cref="Excel.Worksheet"/>.</param>
        public QuizDataLOWrapper(WshListobjPair wshTblNmzPair) : base(wshTblNmzPair)
        {
            _wshTblPair = wshTblNmzPair;
        }
        #endregion

        #region methods
        #region public
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
            QuizDataLOWrapper loWrapper = new QuizDataLOWrapper(_wshTblPair);
            QuizDataLOWrapper.SetStaticHasDataColsPpty();
            QuizDataLOWrapper.SetStaticXLRangePpts();
            if (QuizDataLOWrapper.SortableBLSessions == null)
                SetSortableBindingListPpty();

            colNmbr = 0; // ...to be tested
            byte nmbrTblCols = (byte)loWrapper.XLTable.Range.Columns.Count;

            if (!QuizDataLOWrapper.HasDataCols)
            {
                colNmbr = nmbrTblCols + 1;
            }
            else
            {
                byte nmbrDataCols = (byte)(nmbrTblCols - DataTblNmbrRowLblCols);
                for (int i = 1; i <= nmbrDataCols; i++)
                {
                    if (s > QuizDataLOWrapper.SortableBLSessions[i])
                        colNmbr = i;
                }
                // Now trap for Sess No being smaller than any which are 
                // currently in quiz data table...
                if (colNmbr == 0)
                    colNmbr = nmbrTblCols + 1;
            }

            // Now either insert a column or expand table because our new 
            // Session data will be appended as the right-most column...
            if(colNmbr == nmbrTblCols +1)
            {
                loWrapper.XLTable.Resize(loWrapper.XLTable.Range.Resize[ColumnSize: nmbrTblCols + 1]);
            }
            else
            {
                Excel.Range col = loWrapper.WshParent.Columns[ColumnIndex: colNmbr];
                col.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            }

            // Populate header info...
            int rngCol = colNmbr - DataTblNmbrRowLblCols;
            Excel.Range hdrCell;

            hdrCell = QuizDataLOWrapper.XLRngSessNos.Cells[1, rngCol];
            hdrCell.Value = s.SessionNo;

            hdrCell = QuizDataLOWrapper.XLRngSessDates.Cells[1, rngCol];
            hdrCell.Value = s.QuizDate;

            hdrCell = QuizDataLOWrapper.XLRngCourseWk.Cells[1, rngCol];
            hdrCell.Value = s.CourseWeek;

            hdrCell = QuizDataLOWrapper.XLRngMaxPts.Cells[1, rngCol];
            hdrCell.Value = s.MaxPts;

            hdrCell = QuizDataLOWrapper.XLRngWhichSess.Cells[1, rngCol];
            hdrCell.Value = Session.GetOrdinalNameFromWhichSessEnum(s.WeeklySession);

            hdrCell = loWrapper.XLTable.HeaderRowRange.Cells[1, rngCol];
            hdrCell.Value = s.ColHeaderText;

            // Set the number of table columns field...
            _noLOCols = loWrapper.XLTable.Range.Columns.Count;
        }

        /// <summary>
        /// Populates the 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.QuizDataLOWrapper.SortableBLSessions"/>.
        /// </summary>
        public static void SetSortableBindingListPpty()
        {
            QuizDataLOWrapper.SetStaticHasDataColsPpty();
            QuizDataLOWrapper.SetStaticXLRangePpts();
            QuizDataLOWrapper loWrapper = new QuizDataLOWrapper(_wshTblPair);
            if (QuizDataLOWrapper.HasDataCols)
            {
                if (_srtblBLSessns == null)
                    _srtblBLSessns = new SortableBindingList<Session>();
                byte nmbrTblCols = (byte)loWrapper.XLTable.Range.Columns.Count;
                for(byte i = 1; i <= nmbrTblCols - DataTblNmbrRowLblCols; i++)
                {
                    byte wk = byte.Parse(QuizDataLOWrapper.XLRngCourseWk[1, i].Value);
                    byte mxPts = byte.Parse(QuizDataLOWrapper.XLRngMaxPts[1, i].Value);
                    DateTime qzDt = QuizDataLOWrapper.XLRngSessDates[1, i].Value2;
                    string sessNo = (string)QuizDataLOWrapper.XLRngSessNos[1, i].Value;
                    WkSession sessEnum = 
                        Session.GetWhichSessEnumFromOrdinal(
                            QuizDataLOWrapper.XLRngWhichSess[1, i].Value);
                    _srtblBLSessns.Add(new Session(sessNo, qzDt, mxPts, wk, sessEnum));
                }
            }
        }

        /// <summary>
        /// Populates the fields underlying the properties which get the number (index) 
        /// of columns containing student emails, last names, and first names.
        /// </summary>
        public static void SetStudentInfoColNmbrPptys()
        {
            QuizDataLOWrapper loWrapper = new QuizDataLOWrapper(_wshTblPair);
            _colnoEmls = loWrapper.XLTable.ListColumns["Student ID"].Index;
            _colnoLNms = loWrapper.XLTable.ListColumns["Last Name"].Index;
            _colnoFNms = loWrapper.XLTable.ListColumns["First Name"].Index;
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
                arrxlStudentsToAdd[i - 1, QuizDataLOWrapper.ColNmbrEmails - 1] =
                    newStudentsList[i - 1].EmailAddr;
                arrxlStudentsToAdd[i - 1, QuizDataLOWrapper.ColNmbrLastNms - 1] =
                    newStudentsList[i - 1].LastName;
                arrxlStudentsToAdd[i - 1, QuizDataLOWrapper.ColNmbrFirstNms - 1] =
                    newStudentsList[i - 1].FirstName;
            }

            // Define range where the new student info will be "pasted" (multi-step)...
            QuizDataLOWrapper loWrapper = new QuizDataLOWrapper(_wshTblPair);
            // We are assuming here that email column is still left-most column, and that 
            // last name & firt name columns are, in either order, the next 2 columns...
            Excel.Range rngAddStudents =
                loWrapper.XLTable.ListColumns[_colnoEmls].DataBodyRange;
            rngAddStudents = rngAddStudents.Resize[RowSize: nmbrNewStudents, ColumnSize: 3];
            if (!ThisWbkWrapper.IsVirginWbk)
            {
                // We have to "re-locate" range to just below existing data body range...
                nmbrExistingStudents = loWrapper.XLTable.DataBodyRange.Rows.Count;
                rngAddStudents = rngAddStudents.Offset[RowOffset: nmbrExistingStudents];
            }

            // "Paste" the new student info into the target range...
            rngAddStudents.Value = arrxlStudentsToAdd;

            // Re-define ListObject to encompass the rows of new students...
            Excel.Range newLO = loWrapper.XLTable.HeaderRowRange.Resize[
                RowSize: nmbrExistingStudents + nmbrNewStudents + 1];
            loWrapper.XLTable.Resize(newLO);
        }
        #endregion
        #region private
        private static void SetStaticHasDataColsPpty()
        {
            QuizDataLOWrapper loWrapper = new QuizDataLOWrapper(_wshTblPair);
            if (_hasDataCols)
                return; // ...obviously has already been set
            else
            {
                byte nmbrTblCols = (byte)loWrapper.XLTable.Range.Columns.Count;
                if (nmbrTblCols > DataTblNmbrRowLblCols)
                    _hasDataCols = true;
            }
        }

        private static void SetStaticXLRangePpts()
        {
            if (_staticPptsSet || !_hasDataCols)
                return;
            else
            {
                QuizDataLOWrapper loWrapper = new QuizDataLOWrapper(_wshTblPair);
                byte nmbrTblCols = (byte)loWrapper.XLTable.Range.Columns.Count;
                byte nmbrDataCols = (byte)(nmbrTblCols - DataTblNmbrRowLblCols);
                int rownoSessNos = loWrapper.WshParent.Range["rowSessionNmbr"].Row;
               
                // 1st define Session Nos range as a single cell...
                _rngSessNos = loWrapper.WshParent.Cells[rownoSessNos, DataTblNmbrRowLblCols + 1];
                _rngSessNos = _rngSessNos.Resize[1, nmbrDataCols];
                
                // Define remaining ranges as offsets from _rngSessNos...
                _rngCourseWk = 
                    _rngSessNos.Offset[loWrapper.WshParent.Range["rowCourseWk"].Row - rownoSessNos];
                _rngSessDts = 
                    _rngSessNos.Offset[loWrapper.WshParent.Range["rowSessionDt"].Row - rownoSessNos];
                _rngSessEnums =
                    _rngSessNos.Offset[loWrapper.WshParent.Range["rowSessionEnum"].Row - rownoSessNos];
                _rngMaxPts =
                    _rngSessNos.Offset[loWrapper.WshParent.Range["rowTtlPts"].Row - rownoSessNos];
                _staticPptsSet = true;
            }
        }

        }
#endregion
        #endregion
}
