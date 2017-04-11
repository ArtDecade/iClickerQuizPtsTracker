﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Globalization;
using iClickerQuizPtsTracker.AppExceptions;

namespace iClickerQuizPtsTracker
{
    /// <summary>
    /// Represents a session during which a student takes an iClicker quiz.
    /// </summary>
    public class Session : IEquatable<Session>, IComparable<Session>
    {
        #region fields
        private string _nmbr;
        private DateTime _date;
        private byte _maxPts;
        private string _comboBxText;
        private string _wbkColHdr;
        #endregion

        #region ppts
            #region readOnly
            /// <summary>
            /// The (unique) number of an iClicker quiz Session.
            /// </summary>
            public string SessionNo
            {
                get
                {
                    if (_nmbr.Length == 1)
                        return string.Format($"0{_nmbr}");
                    else
                        return _nmbr;
                }
            }

            /// <summary>
            /// The date of the Session.
            /// </summary>
            public DateTime QuizDate
            {
                get
                { return _date; }
            }

            /// <summary>
            /// The maximum number of points that can be earned on the iClicker 
            /// quiz given during a Session.
            /// </summary>
            public byte MaxPts
            {
                get
                { return _maxPts; }
            }

            /// <summary>
            /// Session information formatted for ComboBox display.
            /// </summary>
            /// <remarks>
            /// Property should return something like "Session 05 - 02/27/2017".
            /// </remarks>
            public string ComboBoxText
            {
                get
                {
                    string fmtdDate = _date.ToString("d", DateTimeFormatInfo.InvariantInfo);
                    return String.Format($"Session {_nmbr} - {fmtdDate}");
                }
            }

            /// <summary>
            /// The column header to be used in the iCLICKERQuizPoints worksheet.
            /// </summary>
            public string ColHeaderText
            {
                get
                {
                    string fmtdDate = _date.ToString("d", DateTimeFormatInfo.InvariantInfo);
                    return string.Format($"Sess {_nmbr} - {fmtdDate}");
                }
            }
            #endregion
            #region readWrite
            /// <summary>
            /// The course week in which the Session is taught.
            /// </summary>
            public byte CourseWeek { get; set; }
            /// <summary>
            /// Which session within the course week.
            /// </summary>
            public WkSession WeeklySession { get; set; }
        #endregion
        #endregion

        #region ctors
        /// <summary>
        /// Instantiates an instance of a <see cref="iClickerQuizPtsTracker.Session"/>.
        /// </summary>
        /// <param name="rawFileHeader">The column header from a raw iClicer data file.</param>
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.InvalidQuizDataHeaderException">The header 
        /// text from the raw data file is not in the expected format.</exception>
        public Session(string rawFileHeader)
        {
            try
            {
                ExtractSessionDataFromColumnHeader(rawFileHeader,
                    out _nmbr, out _date, out _maxPts);
            }
            catch (InvalidQuizDataHeaderException ex)
            {
                throw ex;
            }
            // If necessary add a leading zero to the Session number...
            if (_nmbr.Length == 1)
                _nmbr =  string.Format($"0{_nmbr}");
        }

        /// <summary>
        /// Instantiates an instance of a <see cref="iClickerQuizPtsTracker.Session"/>.
        /// </summary>
        /// <param name="sessNo">The number of the iClicker session.</param>
        /// <param name="sessDate">The date of the session.</param>
        /// <param name="maxPts">The maximum number of points that a student 
        /// can earn from the Session&apos;s iClicker quiz.</param>
        public Session(string sessNo, DateTime sessDate, byte maxPts)
        {
            // This sessNo check SHOULD be unnecessary, but just in case...
            if (sessNo.Length == 1)
                _nmbr = string.Format($"0{sessNo}");
            else
                _nmbr = sessNo;
            _date = sessDate;
            _maxPts = maxPts;
        }
        #endregion

        #region methods
        #region private
        /// <summary>
        /// Obtains the session number, quiz date, and maximum points
        /// from a raw data file data column header.
        /// </summary>
        /// <param name="hdr">A column header from a raw quiz data file.</param>
        /// <param name="sessionNo">An out parameter to capture the session number.</param>
        /// <param name="qzDate">An out parameter to capture the date of the quiz.</param>
        /// <param name="maxPts">An out parameter to capture the maximum points for the quiz.</param>
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.InvalidQuizDataHeaderException">The header 
        /// text from the raw data file is not in the expected format.</exception>
        private void ExtractSessionDataFromColumnHeader(string hdr,
            out string sessionNo, out DateTime qzDate, out byte maxPts)
        {
            try
            {
                hdr = hdr.Replace("Session ",string.Empty);
                hdr = hdr.Replace("Total ", string.Empty);
                // (char)91 = opening bracket (i.e., "[")...
                hdr = hdr.Replace(((char)91).ToString(), string.Empty);
                // (char)93 = closing bracket (i.e., "]")...
                hdr = hdr.Replace(((char)93).ToString(), string.Empty);
                hdr = hdr.Trim();
                // Hdr will now be something like:  "40 5/2/16 2.00"...
                int perSpace = hdr.IndexOf((char)46); // ...(char)46 = period (i.e., ".")
                hdr = hdr.Substring(0, perSpace); // ...remove trailing decimals

                int posSpace1 = hdr.IndexOf((char)32, 1); // ...(char)32 = space (i.e., " ")
                int posSpace2 = hdr.IndexOf((char)32, posSpace1 + 1);

                // Now extract our values...
                sessionNo = hdr.Substring(0, posSpace1);
                if (sessionNo.Length == 1)
                    sessionNo = string.Format($"0{sessionNo}"); // ...add leading zero, if necessary
                qzDate = DateTime.Parse( hdr.Substring(posSpace1 + 1, posSpace2 - posSpace1 - 1));
                maxPts = Byte.Parse(hdr.Substring(posSpace2 + 1));
            }
            catch (Exception e)
            {
                InvalidQuizDataHeaderException ex =
                    new InvalidQuizDataHeaderException(
                        "Failure in ExtractSessionDataFromColumnHeader method.", e);
                ex.HeaderText = hdr;
                throw ex;
            }
        }
        #endregion

        #region public
        /// <summary>
        /// Returns a string that represents the current <see cref="iClickerQuizPtsTracker.Session"/> object.
        /// </summary>
        /// <returns>A string that represents the state of the current <see cref="iClickerQuizPtsTracker.Session"/>.</returns>
        public override string ToString()
        {
            string ms1 = 
                string.Format($"[SessionNo: {SessionNo}; QuizDate: {QuizDate.ToShortDateString()}; ");
            string ms2 = 
                string.Format($"MaxPts: {MaxPts}; ComboBoxText: {ComboBoxText}; ColHeaderText: {ColHeaderText}; ");
            string ms3 =
                string.Format($"CourseWeek: {CourseWeek}; WeeklySession: {WeeklySession.ToString()}");

            string myState = ms1 + ms2 + ms3;
            return myState;
        }

        /// <summary>
        /// Returns a has code for the current <see cref="iClickerQuizPtsTracker.Session"/> object.
        /// </summary>
        /// <returns>A hash code for the current <see cref="iClickerQuizPtsTracker.Session"/>.</returns>
        public override int GetHashCode()
        {
            return this.ToString().GetHashCode();
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current 
        /// <see cref="iClickerQuizPtsTracker.Session"/> object.
        /// </summary>
        /// <param name="obj">The object to compare with the current 
        /// <see cref="iClickerQuizPtsTracker.Session"/></param>
        /// <returns><see langword="true"/> if the specified object is equal to the current 
        /// <see cref="iClickerQuizPtsTracker.Session"/>; otherwise <see langword="false"/>.</returns>
        /// <remarks>The only determinent for whether two <see cref="iClickerQuizPtsTracker.Session"/> objects 
        /// are equal is whether they have the same value for their respective <code>SessionNo</code> 
        /// properties.  Values for all other properties are ignored.  </remarks>
        public override bool Equals(object obj)
        {
            // Run through the generic implementation...
            return this.Equals(obj as Session);
        }

        /// <summary>
        /// Determines whether another <see cref="iClickerQuizPtsTracker.Session"/> 
        /// has the same value for <see cref="iClickerQuizPtsTracker.Session.SessionNo"/> 
        /// as this <see cref="iClickerQuizPtsTracker.Session"/>.
        /// </summary>
        /// <param name="s">The <see cref="iClickerQuizPtsTracker.Session"/> 
        /// to which we want to test for equality.</param>
        /// <returns><see langword="true"/>if 
        /// <list type="bullet">
        /// <item>
        /// <description>both this and <code>s</code> are valid 
        /// <see cref="iClickerQuizPtsTracker.Session"/> references and the values 
        /// for the <see cref="iClickerQuizPtsTracker.Session.SessionNo"/> property 
        /// are equal</description>
        /// </item>
        /// <item>
        /// <description>both this and <code>s</code> are <see langword="null"/> 
        /// references </description>
        /// </item>
        /// </list>
        /// <see langword="false"/>if <list type="bullet">
        /// <item>
        /// <description><code>s</code> is a <see langword="null"/> reference  
        /// but <code>s</code> is a valid 
        /// <see cref="iClickerQuizPtsTracker.Session"/> reference</description>
        /// </item>
        /// <item>
        /// <description><code>s</code> is a valid <see cref="iClickerQuizPtsTracker.Session"/> 
        /// reference but this instance is a <see langword="null"/> reference</description>
        /// </item>
        /// <item>
        /// <description>Both this and <code>s</code> are valid 
        /// <see cref="iClickerQuizPtsTracker.Session"/> references but their values for 
        /// <see cref="iClickerQuizPtsTracker.Session.SessionNo"/> differ</description>
        /// </item>
        /// </list></returns>
        public bool Equals(Session s)
        {
            // If parameter is null, return false...
            if (Object.ReferenceEquals(s, null))
            { return false; }

            // Optimization for a common success case...
            if (Object.ReferenceEquals(this, s))
            { return true; }

            // If run-time types are not exactly the same, return false...
            if (this.GetType() != s.GetType())
            { return false; }

            // Return true if SessionNo pptys match...
            return this.SessionNo == s.SessionNo;
        }

        /// <summary>
        /// Compares the value of the the <code>SessionNo</code> property of this 
        /// instance to value of the same property of a specified 
        /// <see cref="iClickerQuizPtsTracker.Session"/> and returns an indication
        /// of their relative values.
        /// </summary>
        /// <param name="other">A <see cref="iClickerQuizPtsTracker.Session"/> 
        /// instance against which to compare 
        /// <see cref="iClickerQuizPtsTracker.Session.SessionNo"/> values.</param>
        /// <returns><code>0</code> if the <code>SessionNo</code> values are equal.  
        /// Otherwise <code>1</code> if the <code>SessionNo</code> value of this 
        /// instance is higher; <code>-1</code> if the <code>SessionNo</code> 
        /// value of this instance is lower.</returns>
        public int CompareTo(Session other)
        {
            if(Object.ReferenceEquals(this, null) || Object.ReferenceEquals(other,null))
            {
                if (Object.ReferenceEquals(this, null) && Object.ReferenceEquals(other, null))
                    return 0;
                else
                {
                    if (Object.ReferenceEquals(this, null))
                        return -1;
                    else //... other == null
                        return 1;
                }
            }
            else // ...non-null references
            {
                byte sessNmbrThis = byte.Parse(this.SessionNo);
                byte sessNmbrOther = byte.Parse(other.SessionNo);
                return sessNmbrThis.CompareTo(sessNmbrOther);
            }
        }

        /// <summary>
        /// Determines whether two <see cref="iClickerQuizPtsTracker.Session"/> objects 
        /// have the same value for their <code>SessionNo</code> properties.
        /// </summary>
        /// <param name="s1">The first <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <param name="s2">The second <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <returns>
        /// <see langword="true"/> if the values for the <code>SessionNo</code> property are the same for the 
        /// two <see cref="iClickerQuizPtsTracker.Session"/> objects; otherwise <see langword="false"/>.
        /// </returns>
        public static bool operator == (Session s1, Session s2)
        {
            // Check for null on left side...
            if (Object.ReferenceEquals(s1, null))
            {
                if (Object.ReferenceEquals(s2, null))
                {
                    // null == null = true...
                    return true;
                }

                // Only the left side is null...
                return false;
            }
            // Equals handles case of null on right side...
            return s1.Equals(s2);
        }

        /// <summary>
        /// Determines whether two <see cref="iClickerQuizPtsTracker.Session"/> objects 
        /// have different values for their <code>SessionNo</code> properties.
        /// </summary>
        /// <param name="s1">The first <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <param name="s2">The second <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <returns><see langword="true"/> if the values for the <code>SessionNo</code> property differ for the 
        /// two <see cref="iClickerQuizPtsTracker.Session"/> objects; otherwise <see langword="false"/>.</returns>
        public static bool operator != (Session s1, Session s2)
        {
            return !(s1 == s2);
        }

        /// <summary>
        /// Determines whether the value for the <code>SessionNo</code> property of one <see cref="iClickerQuizPtsTracker.Session"/> 
        /// object is less than the value for the <code>SessionNo</code> property of a second 
        /// <see cref="iClickerQuizPtsTracker.Session"/> object.
        /// </summary>
        /// <param name="s1">The first <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <param name="s2">The second <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <returns><see langword="true"/> if the value for the <code>SessionNo</code> property of the first 
        /// <see cref="iClickerQuizPtsTracker.Session"/> objects is less than the value of that same property 
        /// of a second <see cref="iClickerQuizPtsTracker.Session"/> object; otherwise <see langword="false"/>.</returns>
        public static bool operator < (Session s1, Session s2)
        {
            // Nulls are always considered lower in value than an instantiated 
            // object (2 nulls are considered equal)...
            if (Object.ReferenceEquals(s1, null) || Object.ReferenceEquals(s2, null))
            {
                if (Object.ReferenceEquals(s1, null) && Object.ReferenceEquals(s2, null))
                    return false;
                if (Object.ReferenceEquals(s1, null))
                    return true;
                else // s2 == null
                    return false;
            }
            else
            {
                byte sNo1 = byte.Parse(s1.SessionNo);
                byte sNo2 = byte.Parse(s2.SessionNo);
                return sNo1 < sNo2;
            }
        }

        /// <summary>
        /// Determines whether the value for the <code>SessionNo</code> property of one <see cref="iClickerQuizPtsTracker.Session"/> 
        /// object is greater than the value for the <code>SessionNo</code> property of a second 
        /// <see cref="iClickerQuizPtsTracker.Session"/> object.
        /// </summary>
        /// <param name="s1">The first <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <param name="s2">The second <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <returns><see langword="true"/> if the value for the <code>SessionNo</code> property of the first 
        /// <see cref="iClickerQuizPtsTracker.Session"/> objects is greater than the value of that same property 
        /// of a second <see cref="iClickerQuizPtsTracker.Session"/> object; otherwise <see langword="false"/>.</returns>
        public static bool operator > (Session s1, Session s2)
        {
            // Nulls are always considered lower in value than an instantiated 
            // object (2 nulls are considered equal)...
            if (Object.ReferenceEquals(s1, null) || Object.ReferenceEquals(s2, null))
            {
                if (Object.ReferenceEquals(s1, null) && Object.ReferenceEquals(s2, null))
                    return false;
                if (Object.ReferenceEquals(s1, null))
                    return false;
                else // s2 == null
                    return true;
            }
            else
            {
                byte sNo1 = byte.Parse(s1.SessionNo);
                byte sNo2 = byte.Parse(s2.SessionNo);
                return sNo1 > sNo2;
            }
        }

        /// <summary>
        /// Determines whether the value for the <code>SessionNo</code> property of one <see cref="iClickerQuizPtsTracker.Session"/> 
        /// object is less than or equal to the value for the <code>SessionNo</code> property of a second 
        /// <see cref="iClickerQuizPtsTracker.Session"/> object.
        /// </summary>
        /// <param name="s1">The first <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <param name="s2">The second <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <returns><see langword="true"/> if the value for the <code>SessionNo</code> property of the first 
        /// <see cref="iClickerQuizPtsTracker.Session"/> objects is less than or equal to the value of that same property 
        /// of a second <see cref="iClickerQuizPtsTracker.Session"/> object; otherwise <see langword="false"/>.</returns>
        public static bool operator <= (Session s1, Session s2)
        {
            // Nulls are always considered lower in value than an instantiated 
            // object (2 nulls are considered equal)...
            if (Object.ReferenceEquals(s1, null) || Object.ReferenceEquals(s2, null))
            {
                if (Object.ReferenceEquals(s1, null) && Object.ReferenceEquals(s2, null))
                    return true;
                if (Object.ReferenceEquals(s1, null))
                    return true;
                else // s2 == null
                    return false;
            }
            else
            {
                byte sNo1 = byte.Parse(s1.SessionNo);
                byte sNo2 = byte.Parse(s2.SessionNo);
                return sNo1 <= sNo2;
            }
        }

        /// <summary>
        /// Determines whether the value for the <code>SessionNo</code> property of one <see cref="iClickerQuizPtsTracker.Session"/> 
        /// object is greater than or equal to the value for the <code>SessionNo</code> property of a second 
        /// <see cref="iClickerQuizPtsTracker.Session"/> object.
        /// </summary>
        /// <param name="s1">The first <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <param name="s2">The second <see cref="iClickerQuizPtsTracker.Session"/> object being compared.</param>
        /// <returns><see langword="true"/> if the value for the <code>SessionNo</code> property of the first 
        /// <see cref="iClickerQuizPtsTracker.Session"/> objects is greater than or equal to the value of that same property 
        /// of a second <see cref="iClickerQuizPtsTracker.Session"/> object; otherwise <see langword="false"/>.</returns>
        public static bool operator >= (Session s1, Session s2)
        {
            // Nulls are always considered lower in value than an instantiated 
            // object (2 nulls are considered equal)...
            if (Object.ReferenceEquals(s1, null) || Object.ReferenceEquals(s2, null))
            {
                if (Object.ReferenceEquals(s1, null) && Object.ReferenceEquals(s2, null))
                    return true;
                if (Object.ReferenceEquals(s1, null))
                    return false;
                else // s2 == null
                    return true;
            }
            else
            {
                byte sNo1 = byte.Parse(s1.SessionNo);
                byte sNo2 = byte.Parse(s2.SessionNo);
                return sNo1 >= sNo2;
            }
        }
        #endregion
        #endregion
    }
}
