using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iClickerQuizPtsTracker.AppExceptions;
using iClickerQuizPtsTracker.Itfs;
using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPtsTracker
{
    /// <summary>
    /// Provides a mechanism for verifying that named ranges still exist in workbook. 
    /// </summary>
    public class NamedRangeWrapper : INamedRangeWrapping
    {
        /// <summary>
        /// Tests whether a specified workbook-scoped <see cref="Excel.Name"/> both exists 
        /// and refers to a valid <see cref="Excel.Range"/>.
        /// </summary>
        /// <param name="rngNm">The name of the <see cref="Excel.Range"/>.</param>
        /// <returns><c>true</c> if the <see cref="Excel.Name"/> both exists and refers to a valid 
        /// <see cref="Excel.Range"/>; otherwise <c>false</c>.</returns>
        public virtual bool WorkbookScopedRangeExists(string rngNm)
        {
            if (Globals.ThisWorkbook.Names.Count == 0)
                return false;

            bool rngFound = false;
            foreach (Excel.Name n in Globals.ThisWorkbook.Names)
            {
                if (n.Name == rngNm)
                {
                    rngFound = true;
                    break;
                }
            }

            if (!rngFound)
                return false;
            else // ...the named range exists
            {
                // Compiler needs to see that we have, in fact, assigned a value to this variable...
                Excel.Name wbkScopedNm = Globals.ThisWorkbook.Names.Item(rngNm);
               
                // Now see if the named range has a valid reference...
                try
                {
                    Excel.Range r = wbkScopedNm.RefersToRange;
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Tests whether a specified worksheet-scoped <see cref="Excel.Name"/> both exists 
        /// and refers to a valid <see cref="Excel.Range"/>.
        /// </summary>
        /// <param name="wshNm">The name of the parent <see cref="Excel.Worksheet"/>.</param>
        /// <param name="rngNm">The name of the <see cref="Excel.Range"/>.</param>
        /// <returns><c>true</c> if the <see cref="Excel.Name"/> both exists and refers to a valid 
        /// <see cref="Excel.Range"/>; otherwise <c>false</c>.</returns>
        public virtual bool WorksheetScopedRangeExists(string wshNm, string rngNm)
        {
            Excel.Worksheet ws = Globals.ThisWorkbook.Worksheets.Item[wshNm];
            if (ws.Names.Count == 0)
                return false;

            bool rngFound = false;
            foreach(Excel.Name n in ws.Names)
            {
                
                // char(33) == Exclamation pt (i.e., "!")...
                if(n.Name == String.Format($"{wshNm}char(33){rngNm}"))
                {
                    rngFound = true;
                    break;
                }
            }

            if (!rngFound)
                return false;
            else // ...the named range exists
            {
                // Compiler needs to see that we have, in fact, assigned a value to this variable...
                Excel.Name wshScpdNm = ws.Names.Item(rngNm);

                // Now see if the named range has a valid reference...
                try
                {

                    Excel.Range r = wshScpdNm.RefersToRange;
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }
    }
}
