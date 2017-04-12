using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using iClickerQuizPtsTracker.AppExceptions;

namespace iClickerQuizPtsTracker
{
    class AppConfigVals
    {
        #region fields
        #region private
        private static List<String> _badKeys = new List<String>();
        #endregion

        #region public
        public static readonly byte ExtFileColNoStudentEmail;
        public static readonly byte ExtFileColNoStudentName;
        public static readonly byte ExtFileNmbrRowLblCols;
        public static readonly byte DataTblNmbrRowLblCols;

        public static readonly string DataTblColNoID;
        public static readonly string DataTblColNoEmail;
        public static readonly string DataTblColNoLNm;
        public static readonly string DataTblColNoFNm;
        #endregion
        #endregion

        #region pptys
        public static List<String> BadAppConfigKeys
        {
            get
            {
                return _badKeys;
            }
        }
        #endregion

        #region ctor
        static AppConfigVals()
        {
            #region populateReadOnlyFields
            AppSettingsReader ar = new AppSettingsReader();

            // Populate fields from App.Config values.  If any exceptions add the 
            // corrupt key name to _badKeys...
            try
            {
                ExtFileColNoStudentEmail = (byte)ar.GetValue("ColNoEmailXL", typeof(byte));
            }
            catch
            {
                _badKeys.Add("ColNoEmailXL");
            }

            try
            {
                ExtFileColNoStudentName = (byte)ar.GetValue("ColNoStdntNmXL", typeof(byte));
            }
            catch
            {
                _badKeys.Add("ColNoStdntNmXL");
            }

            try
            {
                ExtFileNmbrRowLblCols = (byte)ar.GetValue("NmbrRowLblCols", typeof(Byte));
            }
            catch
            {
                _badKeys.Add("NmbrRowLblCols");
            }

            try
            {
                DataTblNmbrRowLblCols =
                    (byte)ar.GetValue("NmbrNonScoreCols", typeof(byte));
            }
            catch
            {
                _badKeys.Add("NmbrNonScoreCols");
            }

            try
            {
                DataTblColNoID = (string)ar.GetValue("ColID", typeof(string));
            }
            catch
            {
                _badKeys.Add("ColID");
            }

            try
            {
                DataTblColNoEmail = (string)ar.GetValue("ColEmail", typeof(string));
            }
            catch
            {
                _badKeys.Add("ColEmail");
            }

            try
            {
                DataTblColNoFNm = (string)ar.GetValue("ColFN", typeof(string));
            }
            catch
            {
                _badKeys.Add("ColFN");
            }

            try
            {
                DataTblColNoLNm = (string)ar.GetValue("ColLN", typeof(string));
            }
            catch
            {
                _badKeys.Add("ColLN");
            }
#endregion

            // Now check for errors...
            if(_badKeys.Count > 0)
            {
                InalidAppConfigItemException ex = 
                    new InalidAppConfigItemException();
                throw ex;
            }
        }
        #endregion
    }
}
