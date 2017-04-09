using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NUnit.Framework;
using NSubstitute;
using iClickerQuizPtsTracker;
using iClickerQuizPtsTracker.AppExceptions;
using iClickerQuizPtsTracker.ListObjMgmt;
using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts.UnitTests
{
    [TestFixture]
    [Category("ThisWbkWrapperTests")]
    class ThisWbkWrapperTests
    {
        [TestCase("foo")]
        public void VerifyWbkScopedNames_InvalidNames_Throws(string nm)
        {
            var wbw = Substitute.For<ThisWbkWrapper>();

            
        }
    }
}
