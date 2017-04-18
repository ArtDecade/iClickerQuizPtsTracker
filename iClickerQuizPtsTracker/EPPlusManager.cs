using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using iClickerQuizPtsTracker.AppExceptions;
using System.Data;
using System.IO;
using OfficeOpenXml;
using System.ComponentModel;
using static iClickerQuizPtsTracker.AppConfigVals;

namespace iClickerQuizPtsTracker
{
    /// <summary>
    /// Specifies constants defining the results of attempting to open an
    /// Excel file using EPPlus.
    /// </summary>
    public enum ImportResult
    {
        /// <summary>
        /// File opened successfully.
        /// </summary>
        Success = 0,
        /// <summary>
        /// The format of the data in the Excel file is incorrect.
        /// </summary>
        WrongFormat = 1,
        /// <summary>
        /// File a pre-2007 Excel file (i.e., not a *.xlsx file).
        /// </summary>
        NotExcel = 2,
        /// <summary>
        /// File cannot be imported because it is open in another application.
        /// </summary>
        StillOpen = 3
    }

    /// <summary>
    /// Provides a mechanism for utilizing EPPlus to extract data from the 
    /// Excel file containing the raw iClicker quiz points data.
    /// </summary>
    public class EPPlusManager
    {
        #region fields
        private int _lastRow;
        private int _lastCol;
        private string _wbkFullNm;
        private DataTable _dtAllScores;
        private QuizDataParser _hdrParser = new QuizDataParser();
        private Session[] _arrSessions;
        private BindingList<Session> _blistSssnsAll = new BindingList<Session>();
        #endregion

        #region ppts
        /// <summary>
        /// Gets the <see cref="System.Data.DataTable"/> holding the quiz scores
        /// from the raw iClicker data file.
        /// </summary>
        public DataTable RawQuizScoresDataTable
        {
            get
            { return _dtAllScores; }
        }

        /// <summary>
        /// Gets the <see cref="System.ComponentModel.BindingList{Session}"/> containing
        /// all Sessions in the raw iClicker data file.
        /// </summary>
        public BindingList<Session> BListSessionsAll
        {
            get
            { return _blistSssnsAll; }
        }
        #endregion

        #region ctor
        /// <summary>
        /// Creates an instance of the <see cref="iClickerQuizPtsTracker.EPPlusManager"/>
        /// class
        /// </summary>
        /// <param name="wbkFullNm">The full name (i.e., including path) of the
        /// Excel file containing the raw iClicker quiz points data.</param>
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.ReadingExternalWbkException">
        /// The file is either a *.csv and *.xls files (or a different kind 
        /// of file entirely).</exception>
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.InalidAppConfigItemException">
        /// An entry in the <code>appSettings</code> section in the <code>App.config</code> 
        /// file could not be found.
        /// </exception>
        public EPPlusManager(string wbkFullNm)
        {
            if (wbkFullNm.EndsWith("xlsx"))
            {
                _wbkFullNm = wbkFullNm;
            }
            else
            {
                ReadingExternalWbkException ex = new ReadingExternalWbkException();
                ex.ImportResult = ImportResult.NotExcel;
                throw ex;
            }
        }
        #endregion

        #region methods
        /// <summary>
        /// Utilizes EPPlus to create two <see cref="System.Data.DataTable"/>s.  One 
        /// comprises all the data from the quiz data worksheet, the other comprises 
        /// Session number information.
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.ReadingExternalWbkException">
        /// There are any number of problems with the format and/or 
        /// structure of the workbook and/or the worksheet containing the quiz
        /// results data.  The exact nature of the problem is specified in
        /// the exception's message property.</exception>
        /// <exception cref="iClickerQuizPtsTracker.AppExceptions.InvalidQuizDataHeaderException">
        /// A data column header in a raw iClicker data file is not in the expected 
        /// format.  As such, there are problems extracting any of:
        /// <list type="bullet">
        /// <item>session number</item>
        /// <item>quiz/session date</item>
        /// <item>maximum points for the quiz</item>
        /// </list>
        /// </exception>
        /// </summary>
        public virtual void CreateRawQuizDataTable()
        {
            _dtAllScores = new DataTable("RawQuizData");

            using (ExcelPackage p = new ExcelPackage())
            {
                using (FileStream stream = new FileStream(_wbkFullNm, FileMode.Open))
                {
                    // Read the workbook and it's 1st (& presumably only) worksheet...
                    p.Load(stream);
                    ExcelWorksheet ws = p.Workbook.Worksheets[1];

                    /*
                     * 
                     * TRAP FOR PROBLEMS IN THE WORKSHEET...
                     * 
                     */
                    if (ws == null)
                    {
                        ReadingExternalWbkException ex =
                            new ReadingExternalWbkException("No worksheets in the workbook.");
                        ex.ImportResult = ImportResult.WrongFormat;
                        throw ex;
                    }

                    if (!ws.Cells["A1"].Value.ToString().Trim().Equals("Student ID") ||
                        !ws.Cells["B1"].Value.ToString().Trim().Equals("Student Name") ||
                        !ws.Cells["C1"].Value.ToString().Trim().EndsWith("TOTAL"))
                    {
                        string msg = "Incorrect column headings for columns A, B, and/or C";
                        ReadingExternalWbkException ex =
                            new ReadingExternalWbkException(msg);
                        ex.ImportResult = ImportResult.WrongFormat;
                        throw ex;
                    }


                    /*
                     * 
                     * GATHER WORKSHEET DIMENSIONS, TRAPPING FOR MISSING DATA...
                     * 
                     */
                    // Find last col in wsh (header row should always have values)...
                    _lastCol = ws.Dimension.End.Column;
                    while (_lastCol > 1)
                    {
                        ExcelRange c = ws.Cells[1, _lastCol];
                        if (c.Value != null)
                            break;
                        else
                            _lastCol--;
                    }

                    // Trap for no data columns...
                    if (_lastCol <= ExtFileNmbrRowLblCols)
                    {
                        string msg = "There are no columns of quiz data.";
                        ReadingExternalWbkException ex = new ReadingExternalWbkException(msg);
                        ex.ImportResult = ImportResult.WrongFormat;
                        throw ex;
                    }

                    // Find last row of data in wsh.  Student ID column should always have an entry...
                    _lastRow = ws.Dimension.End.Row;
                    while (_lastRow > 1)
                    {
                        ExcelRange c = ws.Cells[_lastRow, ExtFileColNoStudentEmail];
                        if (c.Value != null)
                            break;
                        else
                            _lastRow--;
                    }

                    // Trap for no data rows...
                    if (_lastRow == 1)
                    {
                        string msg = "There are no rows of data.";
                        ReadingExternalWbkException ex = new ReadingExternalWbkException(msg);
                        ex.ImportResult = ImportResult.WrongFormat;
                        throw ex;
                    }


                    /*
                     * 
                     * CREATE COLUMNS & ADD TO DATATABLES...
                     * 
                     */

                    // ALL SCORES DATA TABLE:
                    // Create a primary key column...
                    DataColumn colDataID = new DataColumn(DataTblColNmID, typeof(int));
                    colDataID.AllowDBNull = false;
                    colDataID.AutoIncrement = true;
                    colDataID.ReadOnly = true;
                    colDataID.Unique = true;
                    _dtAllScores.Columns.Add(colDataID);

                    // Create Student ID (email) column...
                    DataColumn colStEml = new DataColumn(DataTblColNmEmail, typeof(string));
                    colStEml.AllowDBNull = false;
                    colStEml.ReadOnly = true;
                    colStEml.Unique = true;
                    _dtAllScores.Columns.Add(colStEml);

                    // Create student Last Name column...
                    DataColumn colLn = new DataColumn(DataTblColNmLNm, typeof(string));
                    colLn.AllowDBNull = false;
                    _dtAllScores.Columns.Add(colLn);

                    // Create & add student First Name column...
                    DataColumn colFn = new DataColumn(DataTblColNmFNm, typeof(string));
                    _dtAllScores.Columns.Add(colFn);

                    // Create & add columns for quiz data...
                    for (int i = ExtFileNmbrRowLblCols + 1; i <= _lastCol; i++)
                    {
                        string rawColHdr = ws.Cells[1, i].Value.ToString().Trim();
                        DataColumn col = new DataColumn(rawColHdr, typeof(byte));
                        col.AllowDBNull = true;
                        try
                        {
                            Session s = new Session(rawColHdr);
                            
                            if (!_blistSssnsAll.Contains(s))
                                _blistSssnsAll.Add(s);
                            else // ...dupe entries
                            {
                                string msg =
                                    string.Format($"Multiple instances of Session {s.SessionNo} are in {_wbkFullNm}.");
                                ReadingExternalWbkException ex =
                                    new ReadingExternalWbkException(msg);
                                ex.ImportResult = ImportResult.WrongFormat;
                                throw ex;
                            }
                            
                            // Set extended properties of column, then add column...
                            col.ExtendedProperties["Session Nmbr"] = s.SessionNo;
                            col.ExtendedProperties["QuizDate"] = s.QuizDate.ToShortDateString();
                            col.ExtendedProperties["MaxQuizPts"] = s.MaxPts.ToString().PadLeft(2, '0');
                            col.ExtendedProperties["ComboBoxLbl"] =
                                string.Format($"Session {s.SessionNo} - {s.QuizDate.ToShortDateString()}");
                            _dtAllScores.Columns.Add(col);
                        }
                        catch
                        {
                            InvalidQuizDataHeaderException ex = new InvalidQuizDataHeaderException();
                            ex.HeaderText = rawColHdr;
                            throw ex;
                        }
                    }


                    /*
                     * 
                     * POPULATE ROWS WITH DATA THEN ADD EACH ROW TO DATATABLE...
                     * 
                     */
                    object stEmail;
                    string studentFullNm;
                    string studentLNm;
                    string studentFNm;

                    // Loop through each data row...
                    for (int rowNo = 2; rowNo <= _lastRow; rowNo++)
                    {
                        stEmail = ws.Cells[rowNo, ExtFileColNoStudentEmail].Value;
                        studentFullNm = ws.Cells[rowNo, ExtFileColNoStudentName].Value.ToString();
                        studentLNm = _hdrParser.ExtractLastNameFromFullName(studentFullNm);
                        studentFNm = _hdrParser.ExtractFirstNameFromFullName(studentFullNm);

                        // Trap for missing email rows...
                        if (stEmail == null)
                        {
                            AddNoEmailStudentToWsh(studentFNm, studentLNm);
                        }
                        else
                        {
                            DataRow r = _dtAllScores.NewRow();
                            // Populate student name & email fields...
                            r[DataTblColNmEmail] = ws.Cells[rowNo, ExtFileColNoStudentEmail].Value.ToString().Trim();
                            r[DataTblColNmLNm] = studentLNm;
                            r[DataTblColNmFNm] = studentFNm;

                            // Loop through each quiz data column...
                            for (int colNo = ExtFileNmbrRowLblCols + 1; colNo <= _lastCol; colNo++)
                            {
                                // Populate quiz data fields...
                                string colNm = ws.Cells[1, colNo].Value.ToString().Trim();
                                object objSc = ws.Cells[rowNo, colNo].Value;
                                if (objSc != null)
                                    r[colNm] = objSc;
                            }
                            _dtAllScores.Rows.Add(r); // ...add row to dataTable
                        }
                    }
                }
            }
        }

        private void AddNoEmailStudentToWsh(string fNm, string lNm)
        {
            Excel.ListObject loNoEml = Globals.WshNoEmail.ListObjects["tblNoEmail"];
            Excel.Range rngFNms = loNoEml.ListColumns["First Name"].DataBodyRange;
            Excel.Range rngLNms = loNoEml.ListColumns["Last Name"].DataBodyRange;

            string existingFNm;
            string existingLNm;
            bool tblHasData = false;
            if (loNoEml.DataBodyRange.Rows.Count > 1)
                tblHasData = true;
            else
            {
                // Yes, we have to cast a cell within a range back to a range type...
                existingFNm = ((Excel.Range)rngFNms[1, 1]).Value;
                existingLNm = ((Excel.Range)rngLNms[1, 1]).Value;

                if ((existingFNm != null) || (existingLNm != null))
                    tblHasData = true;
            }
            if(tblHasData)
            {
                //Excel.Range newRow = loNoEml.InsertRowRange();

            }
            else
            {
                // Just populate the 1 empty data row...
                ((Excel.Range)rngFNms[1, 1]).Value = fNm;
                ((Excel.Range)rngLNms[1, 1]).Value = lNm;
            }


        }
        #endregion
    }
}