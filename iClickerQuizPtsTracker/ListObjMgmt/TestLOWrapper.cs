using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPtsTracker.ListObjMgmt
{
    public static class TestLOWrapper
    {
        #region fields
        private const string LO_NAME = "tblQuizPts";
        private static string _wsNmParent = Globals.WshQuizPts.Name;


        #endregion
        public static LOMgmt LOManager { get; private set; }
        public static Excel.Worksheet ParentWsh { get; private set; }

        static TestLOWrapper()
        {
            LOManager = new LOMgmt();
        }
    }
}
