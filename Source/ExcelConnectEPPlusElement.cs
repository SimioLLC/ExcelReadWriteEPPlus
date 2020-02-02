using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using SimioAPI;
using SimioAPI.Extensions;

using OfficeOpenXml;

namespace ExcelReadWriteEPPlus
{
    class ExcelConnectEPPlusElementDefinition : IElementDefinition
    {
        #region IElementDefinition Members

        /// <summary>
        /// Property returning the full name for this type of element. The name should contain no spaces. 
        /// </summary>
        public string Name
        {
            get { return "ExcelConnectEPPlus"; }
        }

        /// <summary>
        /// Property returning a short description of what the element does.  
        /// </summary>
        public string Description
        {
            get { return "Used with ExcelReadEPPlus and ExcelWriteEPPlus steps.\nThe ExcelConnectEPPlus element may be used in conjunction with the user defined Excel Read and Excel Write steps to read from and write to an Excel spreadsheet."; }
        }

        /// <summary>
        /// Property returning an icon to display for the element in the UI. 
        /// </summary>
        public System.Drawing.Image Icon
        {
            get { return null; }
        }

        /// <summary>
        /// Property returning a unique static GUID for the element.  
        /// </summary>
        public Guid UniqueID
        {
            get { return MY_ID; }
        }

        /// <summary>
        /// We need to use this ID in the element reference property of the Read/Write steps, so we *must* make it public
        /// and unique, so (for example) use Visual Studio Tools > Create Guid > (Registry Format) > Copy to clipboard
        /// and paste it here to create a new unique one.
        /// </summary>
        public static readonly Guid MY_ID = new Guid("{4A0F816D-FB62-4E62-B22A-332D38B461C6}"); //Feb2020/dth

        /// <summary>
        /// Method called that defines the schema (property, display name, description, ...) for the element.
        /// This will show on the Simio UI
        /// </summary>
        public void DefineSchema(IElementSchema schema)
        {
            IPropertyDefinition pd = schema.PropertyDefinitions.AddStringProperty("ExcelWorkbook", String.Empty);
            pd.DisplayName = "Excel Workbook";
            pd.Description = "Path and name to Excel workbook.";
            pd.Required = true;
        }

        /// <summary>
        /// Method called to add a new instance of this element type to a model. 
        /// Returns an instance of the class implementing the IElement interface.
        /// </summary>
        public IElement CreateElement(IElementData data)
        {
            return new ExcelConnectElementEPPlus(data);
        }

        #endregion
    }

    class ExcelConnectElementEPPlus : IElement
    {
        readonly IElementData _data;
        readonly string _readerExcelFileName;
        readonly string _writerExcelFileName;
        readonly List<ExcelWorksheet> _sheets = new List<ExcelWorksheet>();

        ExcelPackage _package;
        ExcelWorkbook _workbook;
        bool _bWroteToWorkbook;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="data"></param>
        public ExcelConnectElementEPPlus(IElementData data)
        {

            _data = data;

            IPropertyReader excelWorkbookProp = _data.Properties.GetProperty("ExcelWorkbook");

            // get filename             
            string fileName = excelWorkbookProp.GetStringValue(_data.ExecutionContext);
            _package = new ExcelPackage(new FileInfo(fileName));

            if (String.IsNullOrEmpty(fileName) == false)
            {
                string fileRoot = null;
                string fileDirectoryName = null;
                string fileExtension = null;

                try
                {
                    fileRoot = System.IO.Path.GetPathRoot(fileName);
                    fileDirectoryName = System.IO.Path.GetDirectoryName(fileName);
                    fileExtension = System.IO.Path.GetExtension(fileName);
                }
                catch (ArgumentException ex)
                {
                    ReportError($"File Root={fileRoot} Directory={fileDirectoryName} Err={ex.Message}", ex.ToString());
                }

                // Get information about the Simio run
                IExecutionInformation info = _data.ExecutionContext.ExecutionInformation;
                string projectFolder = info.ProjectFolder;

                if (String.IsNullOrEmpty(fileDirectoryName) || String.IsNullOrEmpty(fileRoot))
                {
                    fileDirectoryName = projectFolder;
                    fileName = fileDirectoryName + "\\" + fileName;
                }

                _readerExcelFileName = fileName;

                string experimentName = info.ExperimentName;
                if (String.IsNullOrEmpty(experimentName))
                    _writerExcelFileName = fileName;
                else
                {
                    string scenarioName = info.ScenarioName;
                    string replicationNumber = info.ReplicationNumber.ToString();
                    fileName = Path.ChangeExtension(fileName, null);

                    _writerExcelFileName = $"{fileName}_{experimentName}_{scenarioName}_Rep{replicationNumber}{fileExtension}";
                }

            }
        }

        void ReportExcelError(string location, Exception ex, IExecutionContext context)
        {
            context.ExecutionInformation.ReportError($"Excel: Location={location}. Err={ex.Message}");
        }

        /// <summary>
        /// If it has never been initialized, fetch the workbook from our Excel package.
        /// The find the named worksheet within the workbook
        /// </summary>
        /// <param name="pathAndFileName"></param>
        /// <param name="workSheetName"></param>
        /// <returns></returns>
        private ExcelWorksheet ReadSheet(string pathAndFileName, String workSheetName)
        {
            if (_workbook == null)
            {
                try
                {
                    if (_package != null)
                    {
                        _workbook = _package.Workbook;
                    }
                    else
                    {
                        FileInfo excelFile = new FileInfo(pathAndFileName);

                        using (ExcelPackage package = new ExcelPackage(excelFile))
                        {
                            _workbook = package.Workbook;
                            _workbook.Protection.LockWindows = true;
                        }
                    }
                    //// Begin update because we don't need changed till save time
                    //_workbook.BeginUpdate();
                }
                catch (Exception ex)
                {
                    throw new Exception($"There was a problem opening the workbook at '{pathAndFileName}'.\nMessage: {ex.Message}", ex);
                }
            }

            ExcelWorksheet sheet;

            if (_workbook == null)
            {
                throw new Exception($"Worksheet={workSheetName} within Excel={pathAndFileName} Not Found");
            }
            else
            {
                sheet = _workbook.Worksheets[workSheetName];
                if (sheet == null)
                {
                    throw new Exception($"Worksheet={workSheetName} Not Found within the WorkBook of {pathAndFileName}");
                }
                else
                {
                    ExcelAddressBase addressBase = sheet.Dimension;
                    if ( addressBase == null )
                    { // If an empty sheet, then add dummy row and return
                        sheet.InsertRow(1,1);
                    }

                    return sheet;
                }
            }
        }

        /// <summary>
        /// Find the given worksheet
        /// </summary>
        /// <param name="workSheetName"></param>
        /// <param name="rowNumber"></param>
        /// <param name="columnNumber"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public string ReadResults(String workSheetName, Int32 rowNumber, Int32 columnNumber, IExecutionContext context)
        {
            string stringResult = String.Empty;

            if (String.IsNullOrEmpty(_readerExcelFileName))
                ReportFileOpenError("[No file specified]", "writing", "[None]");

            ExcelWorksheet sheet = null;

            foreach (ExcelWorksheet sh in _sheets)
            {
                if (sh.Name == workSheetName)
                {
                    sheet = sh;
                    break;
                }
            }

            if (sheet == null)
            {
                sheet = ReadSheet(_readerExcelFileName, workSheetName);
                _sheets.Add(sheet);
            }

            ExcelRow row = sheet.Row(rowNumber-1); //. .Rows[rowNumber - 1][columnNumber - 1];
            var cell = sheet.GetValue(rowNumber - 1, columnNumber - 1);


            if (cell != null)
            {
                
                var type = cell.GetType();
                switch (type.Name)
                {
                    case "String":
                    case "Decimal":
                    case "Boolean":
                    case "DateTime":
                        break;

                    default:
                        break;
                    
                }
            }
            else
            {
                stringResult = string.Empty;
            }

            return stringResult;
        }

        /// <summary>
        /// Write the double value 
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="rowNumber"></param>
        /// <param name="columnNumber"></param>
        /// <param name="writeDoubleValue"></param>
        /// <param name="writeDateTimeValue"></param>
        /// <param name="writeStringValue"></param>
        /// <param name="context"></param>
        public void WriteResults(String workSheet, Int32 rowNumber, Int32 columnNumber, Double writeDoubleValue, DateTime writeDateTimeValue, String writeStringValue, IExecutionContext context)
        {
            _bWroteToWorkbook = true;

            string stringResult = String.Empty;

            ExcelWorksheet sheet = null;

            foreach (ExcelWorksheet sh in _sheets)
            {
                if (sh.Name == workSheet)
                {
                    sheet = sh;
                    break;
                }
            }

            if (sheet == null)
            {
                sheet = ReadSheet(_readerExcelFileName, workSheet);
                _sheets.Add(sheet);
            }

            ExcelRow er = sheet.Row(rowNumber);
            if (er == null)
            {
                sheet.InsertRow(1, 1);
            }

            var c = sheet.Cells[rowNumber, columnNumber];

            if (writeDoubleValue != System.Double.MinValue)
            {
                c.Value = writeDoubleValue;
            }
            else if (writeDateTimeValue != System.DateTime.MinValue)
            {
                DateTime dt = writeDateTimeValue;
                if (dt.Millisecond >= 995)
                {
                    // Excel stores things as Days from Jan 1, 1900. This can (apparently) result in some values like
                    // 1/7/2016 4:29:59.999 when what was in excel was shown as 1/7/2016 4:30:00, so....
                    // If we are very, very close to the next second, so we'll go to the next second, since 
                    //  the ToString() will simply strip off any sub-second values.
                    dt = dt.AddSeconds(1.0);
                }
                c.Value = dt;
            }
            else
            {
                c.Value = writeStringValue;
            }
        }

        void ReportFileOpenError(string fileName, string action, string exceptionMessage)
        {
            _data.ExecutionContext.ExecutionInformation.ReportError(
                $"Error opening {fileName} for {action}. This may mean the specified file, path or disk does not exist.\n\nInternal exception message: {exceptionMessage}");
        }

        /// <summary>
        /// Report a serious error back up to Simio.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="action"></param>
        /// <param name="exceptionMessage"></param>
        void ReportError(string errorDescription, string exceptionMessage)
        {
            _data.ExecutionContext.ExecutionInformation.ReportError(
                $"Error={errorDescription}. \n\nInternal Details={exceptionMessage}");
        }

        #region IElement Members

        /// <summary>
        /// Method called when the simulation run is initialized.
        /// </summary>
        public void Initialize()
        {
            // No initialization logic needed, we will open the file on the first read or write request
        }

        /// <summary>
        /// Method called when the simulation run is terminating.
        /// </summary>
        public void Shutdown()
        {
            // On shutdown, we need to make sure to close the Excel Connection
            try
            {
                if (_workbook != null && _bWroteToWorkbook)
                {
                    // Now we need changes, so do updates
                    //_workbook.EndUpdate();

                    if (!String.IsNullOrEmpty(_writerExcelFileName))
                    {
                        _package.Save();
                        
                        //_workbook.SaveDocument(_writerExcelFileName);
                    }
                }
                _workbook = null;
                _sheets.Clear();
            }
            catch (Exception e)
            {
                ReportExcelError("Shutdown", e, _data.ExecutionContext);
            }
        }

        #endregion
    }

}
