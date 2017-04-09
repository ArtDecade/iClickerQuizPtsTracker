using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections;


namespace iClickerQuizPtsTracker.Comparers
{
    /// <summary>
    /// Represents a mechanism for comparing two instances of 
    /// <see cref="iClickerQuizPtsTracker.Session"/> based on their 
    /// <code>QuizDate</code> properties.
    /// </summary>
    public class SessionDateComparer : IComparer<Session>
    {
        /// <summary>
        /// Compares two instances of <see cref="iClickerQuizPtsTracker.Session"/> 
        /// based on their <code>QuizDate</code> properties.
        /// </summary>
        /// <param name="s1">First Session to be compared.</param>
        /// <param name="s2">Second Session to be compared.</param>
        /// <returns>1 if <see cref="iClickerQuizPtsTracker.Session.QuizDate"/> 
        /// for the first Session passed in is later than that of the second 
        /// Session passed in. 
        /// <para>-1 if <see cref="iClickerQuizPtsTracker.Session.QuizDate"/> 
        /// for the first Session passed in is earlier than that of the second 
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
                if (s1.QuizDate > s2.QuizDate)
                    return 1;
                if (s1.QuizDate < s2.QuizDate)
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
