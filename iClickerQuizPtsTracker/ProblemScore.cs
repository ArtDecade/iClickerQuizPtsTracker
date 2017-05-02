using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPtsTracker
{
    /// <summary>
    /// Respresents a problematic quiz score.
    /// </summary>
    public class ProblemScore
    {
        #region ctor
        /// <summary>
        /// Instantiates an instance of a <see cref="iClickerQuizPtsTracker.ProblemScore"/>.
        /// </summary>
        /// <param name="stdnt">The <see cref="iClickerQuizPtsTracker.Student"/> with the 
        /// problematic score.</param>
        /// <param name="sess">The <see cref="iClickerQuizPtsTracker.Session"/> during which 
        /// the student posted the problematic quiz score.</param>
        /// <param name="qzScore">The student&#39;s problematic score.</param>
        /// <param name="ignored"><see langword="true"/> if the problematic score has been 
        /// excluded from the student&#39;s total quiz scores in the 
        /// <see cref="Globals.WshQuizPts"/> worksheet.</param>
        /// <remarks>As a rule the <see cref="iClickerQuizPtsTracker.ProblemScore.ScoreIgnored"/> 
        /// property should be set to <see langword="false"/> as it is clearly up the the 
        /// teaching staff to decide how to handle problematic scores.  This property, and 
        /// the &quot;Score Ignored&quot; <see cref="Excel.ListColumn"/> on the appropriate 
        /// worksheets are included to clarify any doubt the user may have about how this 
        /// application handles problematic scores.</remarks>
        public ProblemScore(Student stdnt, Session sess, byte qzScore, bool ignored)
        {
            this.Stdnt = stdnt;
            this.Sess = sess;
            this.QuizScore = qzScore;
            this.ScoreIgnored = ignored;
        }
        #endregion

        #region Pptys
        /// <summary>
        /// The <see cref="iClickerQuizPtsTracker.Student"/> with the problematic 
        /// quiz score.
        /// </summary>
        public Student Stdnt { get; }

        /// <summary>
        /// The <see cref="iClickerQuizPtsTracker.Session"/> during which the 
        /// problematic score was registered.
        /// </summary>
        public Session Sess { get; }

        /// <summary>
        /// The quiz score in question.
        /// </summary>
        public byte QuizScore { get; }

        /// <summary>
        /// <see langword="true"/> if the problematic score is excluded by this application
        /// when calculating the student&#39;s total quiz points for the semester; otherwise 
        /// <see langword="false"/>.
        ///  <remarks>As a rule the this 
        /// property should be set to <see langword="false"/> as it is clearly up the the 
        /// teaching staff to decide how to handle problematic scores.  This property, and 
        /// the &quot;Score Ignored&quot; <see cref="Excel.ListColumn"/> on the appropriate 
        /// worksheets are included to clarify any doubt the user may have about how this 
        /// application handles problematic scores.</remarks>
        /// </summary>
        public bool ScoreIgnored { get; }
#endregion
    }
}
