using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using iClickerQuizPtsTracker.ListObjMgmt;
using iClickerQuizPtsTracker.AppExceptions;
using Excel = Microsoft.Office.Interop.Excel;
using static iClickerQuizPtsTracker.AppConfigVals;

namespace iClickerQuizPtsTracker
{
    /// <summary>
    /// Provides a mechanism for interacting with this workbook in a 
    /// unit-testable manner.
    /// </summary>
    public class ThisWbkWrapper
    {
        #region fields
        private static bool _virginWbk;
        private QuizDataLOWrapper _qdLOWrppr;
        private DblDippersLOWrapper _ddsLOWrppr;
        private StudentsAddedLOWrapper _saLOWrppr;
        private NoEmailLOWrapper _noEmlsWrppr;
        private NamedRangeWrapper _nrWrppr = new NamedRangeWrapper();

        public static readonly WshListobjPair LOInfoQuizData;
        public static readonly WshListobjPair LOInfoDblDpprs;
        public static readonly WshListobjPair LOInfoStdntsAdded;
        public static readonly WshListobjPair LOInfoMssngEmails;
        #endregion

        #region ctor
        static ThisWbkWrapper()
        {
            LOInfoQuizData = new WshListobjPair("tblQuizPts", Globals.WshQuizPts.Name);
            LOInfoDblDpprs = new WshListobjPair("tblDblDippers", Globals.WshDblDpprs.Name);
            LOInfoStdntsAdded = new WshListobjPair("tblFirstQuizDts", Globals.WshStdntsAdded.Name);
            LOInfoMssngEmails = new WshListobjPair("tblNoEmail", Globals.WshNoEmail.Name);
        }

#endregion

        #region ppts
        //// QuizDataListObjMgr QuizData


        /// <summary>
        /// Gets a value indicating whether this workbook is yet populated 
        /// with any student data.
        /// </summary>
        public static bool IsVirginWbk
        {
            get
            { return _virginWbk; }
        }
        #endregion

        #region methods
        /// <summary>
        /// Instantiates the (currently) 2 fields of List Object wrapper classes.
        /// </summary>
        public virtual void InstantiateListObjWrapperClasses()
        {
            // Instantiate quiz data class...
            try
            {
                _qdLOWrppr = new QuizDataLOWrapper(LOInfoQuizData);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
            try
            {
                _qdLOWrppr.SetListObjAndParentWshPpts(LOInfoQuizData);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }

            // Instantiate double dippers class...
            try
            {
                _ddsLOWrppr = new DblDippersLOWrapper(LOInfoDblDpprs);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
            try
            {
                _ddsLOWrppr.SetListObjAndParentWshPpts(LOInfoDblDpprs);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }

            // Instantiate student first-quiz dates class...
            try
            {
                _saLOWrppr = new StudentsAddedLOWrapper(LOInfoStdntsAdded);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
            try
            {
                _saLOWrppr.SetListObjAndParentWshPpts(LOInfoStdntsAdded);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }

            // Instantiate email-less students class...
            try
            {
                _noEmlsWrppr = new NoEmailLOWrapper(LOInfoMssngEmails);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
            try
            {
                _noEmlsWrppr.SetListObjAndParentWshPpts(LOInfoMssngEmails);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Verifies that workbook-scoped named ranges exist, and that they have 
        /// valid range references.
        /// </summary>
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.MissingInvalidNmdRngException">
        /// Caught and rethrown when there are problems with the validity of a 
        /// workbook-scoped named range.</exception>
        public virtual void VerifyWbkScopedNames(params string[] nms)
        {
            for (int i = 0; i < nms.Length; i++)
            {
                string iClkrNm = nms[i];
                if(!_nrWrppr.WorkbookScopedRangeExists(iClkrNm))
                {
                    MissingInvalidNmdRngException ex = 
                        new MissingInvalidNmdRngException(RangeScope.Wkbk, iClkrNm);
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Verifies that worksheet-scoped named ranges exist, and that they have 
        /// valid range references.
        /// </summary>
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.MissingInvalidNmdRngException">
        /// Caught and rethrown when there are problems with the validity of a 
        /// worksheet-scoped named range.</exception>
        public virtual void VerifyWshScopedNames(params string[] nms)
        {
            for (int i = 0; i < nms.Length; i++)
            {
                // Since this is the only sheet holding named ranges...
                string qzDataWshNm = Globals.WshQuizPts.Name; 
                string iClikerNm = nms[i];
                if(!_nrWrppr.WorksheetScopedRangeExists(qzDataWshNm, iClikerNm))
                {
                    MissingInvalidNmdRngException ex =
                        new MissingInvalidNmdRngException(RangeScope.Wksheet, iClikerNm, qzDataWshNm);
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Sets the <see cref="iClickerQuizPtsTracker.ThisWbkWrapper.IsVirginWbk"/> 
        /// property.
        /// </summary>
        /// <remarks>
        /// This method checks the <code>ListObjectHasData</code> property of each 
        /// <see cref="Excel.ListObject"/> in the workbook.
        /// </remarks>
        public virtual void SetVirginWbkProperty()
        {
            if (!QuizDataLOMgr.LOHasData && !_ddsLOWrppr.ListObjectHasData)
                _virginWbk = true;
        }

        /// <summary>
        /// Displays a <see cref="iClickerQuizPtsTracker.FormCourseSemesterQuestionaire"/> to 
        /// the user.
        /// </summary>
        /// <remarks>
        /// The user's input is stored in cells in the upper left-hand portion of 
        /// the <code>iCLICKERQuizPoints</code> worksheet.
        /// <para><b>NOTE:</b>&#8194;The user is only shown this form if and when 
        /// the <see cref="iClickerQuizPtsTracker.ThisWbkWrapper.IsVirginWbk"/> 
        /// property is <see langword="true"/>.</para>
        /// </remarks>
        public virtual void PromptUserForCourseNameAndSemester()
        {
            FormCourseSemesterQuestionaire frm = new FormCourseSemesterQuestionaire();
            frm.ShowDialog();
        }



        #endregion


    }
}
