using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NUnit.Framework;
using NSubstitute;
using iClickerQuizPtsTracker;
using iClickerQuizPtsTracker.AppExceptions;

namespace iClickerQuizPts.UnitTests
{
    [TestFixture]
    [Category("SessionComparer")]
    public class SessionComparerTests
    {


        public void VerifyDateComparer_1GT2_ReturnsPos1(string dt1, string dt2)
        {

        }

        public void VerifyDateComparer_1LT2_ReturnsNeg1(string dt1, string dt2)
        {

        }

        public void VerifyDateComparer_1EQ2_RtnsZero(string dt1, string dt2)
        {

        }

        public void VerifyDateComparer_Sess1Null_Throws(string dt2)
        {

        }

        public void VerifyDateComparer_Sess02Null_Throws(string dt1)
        {

        }

        public void VerifyDateComparer_BothSessnsNull_Throws()
        {

        }





    }
}