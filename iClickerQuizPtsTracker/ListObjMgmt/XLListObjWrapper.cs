using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using iClickerQuizPtsTracker.AppExceptions;

namespace iClickerQuizPtsTracker.ListObjMgmt
{
    /// <summary>
    /// Provides a mechanism for interacting with the workbook's <see cref="Excel.ListObjects"/>.
    /// </summary>
    public abstract class XLListObjWrapper
    {
        #region Fields
        private Excel.Worksheet _ws = null;
        private Excel.ListObject _lo = null;
        private bool _wshAndListObjIntegrityVerified = false;
        private bool _listObjHasData = false;
        #endregion

        #region ppts
        /// <summary>
        /// Gets the underlying Excel table.
        /// </summary>
        public Excel.ListObject XLTable
        {
            get { return _lo; }
        }

        /// <summary>
        /// Gets the parent worksheet of the Excel table.
        /// </summary>
        public Excel.Worksheet WshParent
        {
            get { return _ws; }
        }

        /// <summary>
        /// Gets a value indicating whether the underlying 
        /// <see cref="Excel.ListObject"/> has yet been populated 
        /// with any data.
        /// </summary>
        public virtual bool ListObjectHasData
        {
            get
            { return _listObjHasData; }
        }

        /// <summary>
        /// Gets a value indicating whether the existence and names of the underlying
        /// <see cref="Excel.Worksheet"/> and <see cref="Excel.ListObject"/> have been
        /// verified.
        /// </summary>
        /// <remarks>
        /// Adding this property gives us a mechanism for ensuring that the 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.XLListObjWrapper.SetListObjAndParentWshPpts"/> 
        /// method has been called prior to any other class method being called.  (Again, the 
        /// requirements of unit testing prevent us from calling the 
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.XLListObjWrapper.SetListObjAndParentWshPpts"/> 
        /// method from within the <see cref="iClickerQuizPtsTracker.ListObjMgmt.XLListObjWrapper"/> 
        /// constructor.
        /// </remarks>
        public virtual bool UnderlyingWshAndListObjVerified
        {
            get
            { return _wshAndListObjIntegrityVerified; }
        }
        #endregion

        #region ctor
        /// <summary>
        /// Initializes a new instance of the 
        /// class <see cref="iClickerQuizPtsTracker.ListObjMgmt.XLListObjWrapper"/>.
        /// </summary>
        /// <param name="wshTblNmzPair">The properties of this <see langword="struct"/> 
        /// should be populated with the name of the <see cref="Excel.ListObject"/> 
        /// and the name of the parent <see cref="Excel.Worksheet"/>.</param>
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.InvalidWshListObjPairException"> thrown
        /// whenever the <see cref="iClickerQuizPtsTracker.WshListobjPair.WshNm"/> property 
        /// or the the <see cref="iClickerQuizPtsTracker.WshListobjPair.ListObjName"/> property 
        /// has not been populated.  (In other words, <i>both</i> properties must contain non-empty, 
        /// non-null values.)</exception>
        protected XLListObjWrapper(WshListobjPair wshTblNmzPair)
        {
            // Trap to ensure that constructor parameter has been populated with both
            // a wsh name and a ListObject name...
            if (!wshTblNmzPair.PptsSet)
            {
                InvalidWshListObjPairException ex = new InvalidWshListObjPairException();
                ex.WshListObjPair = wshTblNmzPair;
                throw ex;
            }
        }
        #endregion

        #region methods
        /// <summary>
        /// Sets <list type="bullet">
        /// <item>parent <see cref="Excel.Worksheet"/> of <see cref="Excel.ListObject"/></item>
        /// <item><see cref="Excel.ListObject"/> itself</item>
        /// <item><see cref="iClickerQuizPtsTracker.ListObjMgmt.XLListObjWrapper.DoesListObjHaveData"/> property</item>
        /// </list>
        /// </summary>
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.MissingWorksheetException"> thrown when the 
        /// parent <see cref="Excel.Worksheet"/> has either been renamed or deleted.</exception>
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.MissingListObjectException"> thrown when
        /// the <see cref="Excel.ListObject"/> has either been renamed or deleted.</exception>
        /// <remarks>It would be more efficient to call this method from within the class 
        /// constructor.  However, doing some complicates unit testing.</remarks>
        public virtual void SetListObjAndParentWshPpts(WshListobjPair wshTblNmzPair)
        {
            if (!DoesParentWshExist(wshTblNmzPair))
            {
                MissingWorksheetException ex = new MissingWorksheetException();
                ex.WshListObjPair = wshTblNmzPair;
                throw ex;
            }
            else
                _ws = Globals.ThisWorkbook.Worksheets[wshTblNmzPair.WshNm];

            if (!DoesListObjExist(wshTblNmzPair))
            {
                MissingListObjectException ex = new MissingListObjectException();
                ex.WshListObjPair = wshTblNmzPair;
                throw ex;
            }
            else
                _lo = _ws.ListObjects[wshTblNmzPair.ListObjName];

            // Set fields...
            _wshAndListObjIntegrityVerified = true;
            _listObjHasData = DoesListObjHaveData();
        }

        public virtual Excel.Worksheet SetParentWsh(WshListobjPair pr)
        {
            if (!DoesParentWshExist(pr))
            {
                MissingWorksheetException ex = new MissingWorksheetException();
                ex.WshListObjPair = pr;
                throw ex;
            }
            else
                return Globals.ThisWorkbook.Worksheets[pr.WshNm];
        }

        public virtual Excel.ListObject SetListObject(WshListobjPair wsloPair)
        {
            if (!DoesListObjExist(wsloPair))
            {
                MissingListObjectException ex = new MissingListObjectException();
                ex.WshListObjPair = wsloPair;
                throw ex;
            }
            else
                return 
        }

        /// <summary>
        /// Determines whether the parent <see cref="Excel.Worksheet"/> of 
        /// the <see cref="Excel.ListObject"/> exists.
        /// </summary>
        /// <returns><c>true</c> if the <see cref="Excel.Worksheet"/> exists; 
        /// otherwise <c>false</c>.</returns>
        public virtual bool DoesParentWshExist(WshListobjPair wsLOPr)
        {
            foreach(Excel.Worksheet ws in Globals.ThisWorkbook.Worksheets)
            {
                if(ws.Name == wsLOPr.WshNm)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Determines whether the the <see cref="Excel.ListObject"/> exists.
        /// </summary>
        /// <returns><c>true</c> if the <see cref="Excel.ListObject"/> exists; 
        /// otherwise <c>false</c>.</returns>
        public virtual bool DoesListObjExist(WshListobjPair wsLOPr)
        {
            if (_ws.ListObjects.Count == 0)
                return false;
            else
            {
                foreach(Excel.ListObject tbl in _ws.ListObjects)
                {
                    if(tbl.Name == wsLOPr.ListObjName)
                    {
                        return true;
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// Determines whether the the <see cref="Excel.ListObject"/> has yet 
        /// been populated with any data.
        /// </summary>
        /// <returns><c>true</c> if the <see cref="Excel.ListObject"/> 
        /// contains data; otherwise <c>false</c>.</returns>
        protected virtual bool DoesListObjHaveData()
        {
            // Now see if there are data...
            if (this.XLTable.ListRows.Count > 1)
                return true;
            else 
            {
                // DataBodyRange is only 1 row...
                foreach(Excel.Range c in this.XLTable.DataBodyRange)
                {
                    if (c.Value != null)
                        return true;
                }
                return false;
            }
        }
        #endregion
    }
}
