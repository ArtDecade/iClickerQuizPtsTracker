using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iClickerQuizPtsTracker.ListObjMgmt
{
    /// <summary>
    /// Provides a mechanism for interacting with the workbook's <see cref="Excel.ListObject"/> 
    /// of students whose names show up in raw iClicker quiz data files even though there
    /// is no email shown for the student.
    /// </summary>
    class NoEmailLOWrapper : XLListObjWrapper
    {
        /// <summary>
        /// Initializes a new instance of the 
        /// class <see cref="iClickerQuizPtsTracker.ListObjMgmt.NoEmailLOWrapper"/>.
        /// </summary>
        /// <param name="wshTblNmzPair">The properties of this <see langword="struct"/> 
        /// should be populated with the name of the <see cref="Excel.ListObject"/> 
        /// containing email-less students and the name 
        /// of the parent <see cref="Excel.Worksheet"/>.</param>
        public NoEmailLOWrapper(WshListobjPair wshTblNmzPair) : base(wshTblNmzPair)
        {
        }
    }
}
