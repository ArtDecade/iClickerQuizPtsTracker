using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iClickerQuizPtsTracker
{
    /// <summary>
    /// Represents a student in the iClicker quiz data.
    /// </summary>
    public class Student
    {
        #region ctors
        /// <summary>
        /// Instantiates an instance of a <see cref="iClickerQuizPtsTracker.Student"/>.
        /// </summary>
        /// <param name="email">A student&#39;s email address.</param>
        /// <param name="comboNm">A student&#39;s full name, in the form &quot;Doe, John&quot;.</param>
        public Student(string email, string comboNm)
        {
            EmailAddr = email.Trim();
            comboNm = comboNm.Trim();
            if (comboNm == ((char)44).ToString()) // ...(char)44 = comma
                return;
            if (comboNm != string.Empty)
            {
                LastName = QuizDataParser.ExtractLastNameFromFullName(comboNm);
                FirstName = QuizDataParser.ExtractFirstNameFromFullName(comboNm);
            }
        }

        /// <summary>
        /// Instantiates an instance of a <see cref="iClickerQuizPtsTracker.Student"/>.
        /// </summary>
        /// <param name="email">A student&#39;s email address.</param>
        /// <param name="lNm">A student&#39;s last name.</param>
        /// <param name="fNm">A student&#39;s first name.</param>
        public Student(string email, string lNm, string fNm)
        {
            EmailAddr = email.Trim();
            LastName = lNm?.Trim();
            FirstName = fNm?.Trim();
        }
        #endregion

        #region pptys
        /// <summary>
        /// Gets a student&#39;s email address.
        /// </summary>
        public string EmailAddr { get; }
        /// <summary>
        /// Gets a student&#39;s last name.
        /// </summary>
        public string LastName { get; }
        /// <summary>
        /// Gets a student&#39;s first name.
        /// </summary>
        public string FirstName { get; }
        #endregion


        #region methods
        /// <summary>
        /// Returns a string that represents the current <see cref="iClickerQuizPtsTracker.Student"/> object.
        /// </summary>
        /// <returns>A string that represents the state of the current <see cref="iClickerQuizPtsTracker.Student"/>.</returns>
        public override string ToString()
        {
            return string.Format($"[EmailAddr: {EmailAddr}; LastName: {LastName}; FirstName: {FirstName}]");
        }

        /// <summary>
        /// Returns a hash code for the current <see cref="iClickerQuizPtsTracker.Student"/> object.
        /// </summary>
        /// <returns>A hash code for the current <see cref="iClickerQuizPtsTracker.Student"/>.</returns>
        public override int GetHashCode()
        {
            return this.ToString().GetHashCode();
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current 
        /// <see cref="iClickerQuizPtsTracker.Student"/> object.
        /// </summary>
        /// <param name="obj">The object to compare with the current 
        /// <see cref="iClickerQuizPtsTracker.Student"/></param>
        /// <returns><see langword="true"/> if the specified object is equal to the current 
        /// <see cref="iClickerQuizPtsTracker.Student"/>; otherwise <see langword="false"/>.</returns>
        /// <remarks>The only determinent for whether two <see cref="iClickerQuizPtsTracker.Student"/> objects 
        /// are equal is whether they have the same value for their respective <code>EmailAddr</code> 
        /// properties.  Values for all other properties are ignored.  </remarks>
        public override bool Equals(object obj)
        {
            // Run through the generic implementation...
            return this.Equals(obj as Student);
        }

        /// <summary>
        /// Determines whether another <see cref="iClickerQuizPtsTracker.Student"/> 
        /// has the same value for <see cref="iClickerQuizPtsTracker.Student.EmailAddr"/> 
        /// as this <see cref="iClickerQuizPtsTracker.Student"/>.
        /// </summary>
        /// <param name="s">The <see cref="iClickerQuizPtsTracker.Student"/> 
        /// to which we want to test for equality.</param>
        /// <returns><see langword="true"/>if 
        /// <list type="bullet">
        /// <item>
        /// <description>both this and <code>s</code> are valid 
        /// <see cref="iClickerQuizPtsTracker.Student"/> references and the values 
        /// for the <see cref="iClickerQuizPtsTracker.Student.EmailAddr"/> property 
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
        /// <see cref="iClickerQuizPtsTracker.Student"/> reference</description>
        /// </item>
        /// <item>
        /// <description><code>s</code> is a valid <see cref="iClickerQuizPtsTracker.Student"/> 
        /// reference but this instance is a <see langword="null"/> reference</description>
        /// </item>
        /// <item>
        /// <description>Both this and <code>s</code> are valid 
        /// <see cref="iClickerQuizPtsTracker.Student"/> references but their values for 
        /// <see cref="iClickerQuizPtsTracker.Student.EmailAddr"/> differ</description>
        /// </item>
        /// </list></returns>
        public bool Equals(Student s)
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

            // Return true if EmailAddr pptys match...
            return this.EmailAddr == s.EmailAddr;
        }

        /// <summary>
        /// Determines whether two <see cref="iClickerQuizPtsTracker.Student"/> objects 
        /// have the same value for their <code>EmailAddr</code> properties.
        /// </summary>
        /// <param name="lhs">The first <see cref="iClickerQuizPtsTracker.Student"/> object being compared.</param>
        /// <param name="rhs">The second <see cref="iClickerQuizPtsTracker.Student"/> object being compared.</param>
        /// <returns>
        /// <see langword="true"/> if the values for the <code>EmailAddr</code> property are the same for the 
        /// two <see cref="iClickerQuizPtsTracker.Student"/> objects; otherwise <see langword="false"/>.
        /// </returns>
        public static bool operator ==(Student lhs, Student rhs)
        {
            // Check for null on left side...
            if (Object.ReferenceEquals(lhs, null))
            {
                if (Object.ReferenceEquals(rhs, null))
                {
                    // null == null = true...
                    return true;
                }

                // Only the left side is null...
                return false;
            }
            // Equals handles case of null on right side...
            return lhs.Equals(rhs);
        }

        /// <summary>
        /// Determines whether two <see cref="iClickerQuizPtsTracker.Student"/> objects 
        /// have different values for their <code>EmailAddr</code> properties.
        /// </summary>
        /// <param name="lhs">The first <see cref="iClickerQuizPtsTracker.Student"/> object being compared.</param>
        /// <param name="rhs">The second <see cref="iClickerQuizPtsTracker.Student"/> object being compared.</param>
        /// <returns><see langword="true"/> if the values for the <code>EmailAddr</code> property differ for the 
        /// two <see cref="iClickerQuizPtsTracker.Student"/> objects; otherwise <see langword="false"/>.</returns>
        public static bool operator !=(Student lhs, Student rhs)
        {
            return !(lhs == rhs);
        }
        #endregion


    }
}
