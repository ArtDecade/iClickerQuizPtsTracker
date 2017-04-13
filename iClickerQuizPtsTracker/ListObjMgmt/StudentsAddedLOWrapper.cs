using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iClickerQuizPtsTracker.ListObjMgmt
{
    /// <summary>
    /// Provides a mechanism for interacting with the workbook's 
    /// <see cref="Excel.ListObject"/> of students-added-by-date.
    /// </summary>
    public class StudentsAddedLOWrapper : XLListObjWrapper
    {
        /// <summary>
        /// Initializes a new instance of the class
        /// <see cref="iClickerQuizPtsTracker.ListObjMgmt.StudentsAddedLOWrapper"/>.
        /// </summary>
        /// <param name="wshTblNmzPair">The properties of this <see langword="struct"/> 
        /// should be populated with the name of the <see cref="Excel.ListObject"/> 
        /// containing the information about when students were added to 
        /// the iClicker quiz data worksheet and the name 
        /// of the parent <see cref="Excel.Worksheet"/>.</param>
        public StudentsAddedLOWrapper(WshListobjPair wshTblNmzPair) : 
            base(wshTblNmzPair)
        {
        }
    }
}
