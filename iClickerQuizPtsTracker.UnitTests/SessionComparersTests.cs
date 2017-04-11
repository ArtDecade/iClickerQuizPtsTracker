using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NUnit.Framework;
using NSubstitute;
using iClickerQuizPtsTracker;
using iClickerQuizPtsTracker.Comparers;
using iClickerQuizPtsTracker.AppExceptions;

namespace iClickerQuizPts.UnitTests
{
    [TestFixture]
    [Category("SessionComparer")]
    public class SessionComparerTests
    {
        const string SESS_NO = "01";
        const byte MAX_PTS = 2;
        const int S1_GT_S2 = 1;
        const int S1_LT_S2 = -1;
        const int S1_EQ_S2 = 0;

        [TestCase("04/05/2017","04/05/2012")]
        [TestCase("4/5/2017", "4/5/2012")]
        [TestCase("04/05/2017", "04/04/2017")]
        public void VerifyDateComparer_1GT2_ReturnsPos1(string dt1, string dt2)
        {
            DateTime qzDt01 = DateTime.Parse(dt1);
            DateTime qzDt02 = DateTime.Parse(dt2);
            Session s1 = new Session(SESS_NO, qzDt01, MAX_PTS);
            Session s2 = new Session(SESS_NO, qzDt02, MAX_PTS);
            SessionDateComparer comparerDt = new SessionDateComparer();

            int comprsn = comparerDt.Compare(s1, s2);

            Assert.AreEqual(S1_GT_S2, comprsn);
        }

        [TestCase("04/05/2017", "04/05/2018")]
        [TestCase("4/5/2017", "4/5/2018")]
        [TestCase("04/05/2017", "04/06/2017")]
        public void VerifyDateComparer_1LT2_ReturnsNeg1(string dt1, string dt2)
        {
            DateTime qzDt01 = DateTime.Parse(dt1);
            DateTime qzDt02 = DateTime.Parse(dt2);
            Session s1 = new Session(SESS_NO, qzDt01, MAX_PTS);
            Session s2 = new Session(SESS_NO, qzDt02, MAX_PTS);
            SessionDateComparer comparerDt = new SessionDateComparer();

            int comprsn = comparerDt.Compare(s1, s2);

            Assert.AreEqual(S1_LT_S2, comprsn);
        }

        [TestCase("02/02/1993", "2/2/1993")]
        [TestCase("2/2/93", "2/2/1993")]
        [TestCase("02/02/1993", "2/2/93")]
        public void VerifyDateComparer_1EQ2_RtnsZero(string dt1, string dt2)
        {
            DateTime qzDt01 = DateTime.Parse(dt1);
            DateTime qzDt02 = DateTime.Parse(dt2);
            Session s1 = new Session(SESS_NO, qzDt01, MAX_PTS);
            Session s2 = new Session(SESS_NO, qzDt02, MAX_PTS);
            SessionDateComparer comparerDt = new SessionDateComparer();

            int comprsn = comparerDt.Compare(s1, s2);

            Assert.AreEqual(S1_EQ_S2, comprsn);
        }

        [TestCase("04/05/2012")]
        public void VerifyDateComparer_Sess01Null_RtnsNeg1(string dt2)
        {
            DateTime qzDt02 = DateTime.Parse(dt2);
            Session s1 = null;
            Session s2 = new Session(SESS_NO, qzDt02, MAX_PTS);
            SessionDateComparer comparerDt = new SessionDateComparer();

            int comprsn = comparerDt.Compare(s1, s2);

            Assert.AreEqual(-1, comprsn);
        }

        [TestCase("04/05/2012")]
        public void VerifyDateComparer_Sess02Null_RtnsPos1(string dt1)
        {
            DateTime qzDt01 = DateTime.Parse(dt1);
            Session s1 = new Session(SESS_NO, qzDt01, MAX_PTS);
            Session s2 = null;
            SessionDateComparer comparerDt = new SessionDateComparer();

            int comprsn = comparerDt.Compare(s1, s2);

            Assert.AreEqual(1, comprsn);
        }

        [TestCase]
        public void VerifyDateComparer_BothSessnsNull_Rtns0()
        {
            Session s1 = null;
            Session s2 = null;
            SessionDateComparer comparerDt = new SessionDateComparer();

            int comprsn = comparerDt.Compare(s1, s2);

            Assert.Zero(comprsn);
        }
    }
}