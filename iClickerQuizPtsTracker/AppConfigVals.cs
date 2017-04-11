using System;

using System.Configuration;
using System.Collections.Generic;

public static class AppConfigVals
{
    private static byte _xlcolnoRawEmail;
    private static byte _xlcolnoRawNm;
    private static byte _xlcolnoRawDataBegins;
    private static byte _dataNmbrRowLblCols;

    private static string _datahdrID;
    private static string _datahdrEmail;
    private static string _datahdrLN;
    private static string _datahdrFN;

    private static List<String> _badKeys = new List<String>();

	public static AppConfigVals()
	{
        ReadAppConfigDataIntoFields();
    }

    private static void ReadAppConfigDataIntoFields()
    {
        AppSettingsReader ar = new AppSettingsReader();

        // Populate fields from App.Config values.  If any exceptions add the 
        // corrupt key name to _badKeys...
        try
        {
            _xlcolnoRawEmail = (byte)ar.GetValue("ColNoEmailXL", typeof(byte));
        }
        catch
        {
            _badKeys.Add("ColNoEmailXL");
        }

        try
        {
            _xlcolnoRawNm = (byte)ar.GetValue("ColNoStdntNmXL", typeof(byte));
        }
        catch
        {
            _badKeys.Add("ColNoStdntNmXL");
        }

        try
        {
            _dataNmbrRowLblCols =
                (byte)ar.GetValue("NmbrNonScoreCols", typeof(byte));
        }
        catch
        {
            _badKeys.Add("NmbrNonScoreCols");
        }

        try
        {
            _datahdrID = (string)ar.GetValue("ColID", typeof(string));
        }
        catch
        {
            _badKeys.Add("ColID");
        }

        try
        {
            _datahdrEmail = (string)ar.GetValue("ColEmail", typeof(string));
        }
        catch
        {
            _badKeys.Add("ColEmail");
        }

        try
        {
            _datahdrFN = (string)ar.GetValue("ColFN", typeof(string));
        }
        catch
        {
            _badKeys.Add("ColFN");
        }

        try
        {
            _datahdrLN = (string)ar.GetValue("ColLN", typeof(string));
        }
        catch
        {
            _badKeys.Add("ColLN");
        }
    }
}
