using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iClickerQuizPtsTracker.Comparers
{
    /// <summary>
    /// Provides a mechanism for comparing two instances of 
    /// <see cref="iClickerQuizPtsTracker.Session"/> based on their 
    /// <code>SessionNo</code> properties.
    /// </summary>
    public class SessionNmbrComparer : IComparer<Session>
    {
        /// <summary>
        /// Compares two instances of <see cref="iClickerQuizPtsTracker.Session"/> 
        /// based on their <code>SessionNo</code> properties.
        /// </summary>
        /// <param name="s1">First Session to be compared.</param>
        /// <param name="s2">Second Session to be compared.</param>
        /// <returns>1 if <see cref="iClickerQuizPtsTracker.Session.SessionNo"/> 
        /// for the first Session passed in is greater than that of the second 
        /// Session passed in. 
        /// <para>-1 if <see cref="iClickerQuizPtsTracker.Session.SessionNo"/> 
        /// for the first Session passed in is lower than that of the second 
        /// object Session in. </para>
        /// <para>Otherwise 0.</para>
        /// </returns>
        /// <exception cref="ArgumentException">Either or both parameters 
        /// is/are a <see langword="null"/> reference.
        /// </exception>
        public int Compare(Session s1, Session s2)
        {
            if (s1 != null && s2 != null)
            {
                if (int.Parse(s1.SessionNo) > int.Parse(s2.SessionNo))
                    return 1;
                if (int.Parse(s1.SessionNo) < int.Parse(s2.SessionNo))
                    return -1;
                else
                    return 0;
            }
            else
            {
                if (s1 == null)
                    throw new ArgumentNullException("s1");
                else
                    throw new ArgumentNullException("s2");
            }
        }
    }
}
