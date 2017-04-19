using System;
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
        private static WshListobjPair _wshTblPair;

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

        public static void AddEmptyDataColumnWithHeaderInfo(Session s, out byte colNmbr)
        {
            QuizDataLOWrapper loWrapper = new QuizDataLOWrapper(_wshTblPair);
            bool hasDataCols = false;
            byte nmbrTblCols = (byte)loWrapper.XLTable.Range.Columns.Count;
            if (nmbrTblCols > DataTblNmbrRowLblCols)
                hasDataCols = true;
            if(!hasDataCols)
            {
                // Don't understand why intellisense demands this cast, but it does...
                colNmbr = (byte)(nmbrTblCols + 1);
            }
            else
            {
                byte nmbrDataCols = (byte)(nmbrTblCols - DataTblNmbrRowLblCols);
                int rownoSessNos = loWrapper.WshParent.Range["rowSessionNmbr"].Row;
                // 2-step definition...
                Excel.Range rngSessNos = 
                    loWrapper.WshParent.Cells[rownoSessNos, DataTblNmbrRowLblCols + 1];
                rngSessNos = rngSessNos.Resize[1, nmbrDataCols];
                System.Array arrxlSessNos = (System.Array)rngSessNos.Value;
                string[,] x = rngSessNos.Value;
                List<string> sessNos = new List<string>();
                for (int i = 1; i <= nmbrDataCols; i++)
                    sessNos.Add((string)x[i, 1]);
                System.Array[,] y = rngSessNos.Value;
                MsgBoxGenerator.SetInvalidHdrMsg(y[1, 2].ToString());

                colNmbr = 0;
            }




            //loWrapper.XLTable.Resize(loWrapper.XLTable.Range.Resize[, DataTblNmbrRowLblCols + 1]);
        }
    }
}
