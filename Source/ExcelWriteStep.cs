using System;
using System.Globalization;
using SimioAPI;
using SimioAPI.Extensions;

namespace ExcelReadWriteEPPlus
{
    class ExcelWriteStepDefinition : IStepDefinition
    {
        #region IStepDefinition Members

        /// <summary>
        /// Property returning the full name for this type of step. The name should contain no spaces. 
        /// </summary>
        public string Name
        {
            get { return "ExcelWriteEPPlus"; }
        }

        /// <summary>
        /// Property returning a short description of what the step does.  
        /// </summary>
        public string Description
        {
            get { return "The Excel Write step may be used to write values to an Excel worksheet using EPPlus."; }
        }

        /// <summary>
        /// Property returning an icon to display for the step in the UI. 
        /// </summary>
        public System.Drawing.Image Icon
        {
            get { return null; }
        }

        /// <summary>
        /// Property returning a unique static GUID for the step.  
        /// </summary>
        public Guid UniqueID
        {
            get { return MY_ID; }
        }
        static readonly Guid MY_ID = new Guid("{E49512CB-D080-4222-A11A-0AB41016A1CA}"); // Feb2020/dth

        /// <summary>
        /// Property returning the number of exits out of the step. Can return either 1 or 2. 
        /// </summary>
        public int NumberOfExits
        {
            get { return 1; }
        }

        /// <summary>
        /// Method called that defines the property schema for the step.
        /// </summary>
        public void DefineSchema(IPropertyDefinitions schema)
        {
            IPropertyDefinition pd;

            // Reference to the excel to write to
            pd = schema.AddElementProperty("ExcelConnectEPPlus", ExcelConnectEPPlusElementDefinition.MY_ID);

            // And a format specifier
            pd = schema.AddStringProperty("Worksheet", String.Empty);
            pd.Description = "Worksheet Property";
            pd.Required = true;

            pd = schema.AddExpressionProperty("Row", "1");
            pd.Description = "Row Property";
            pd.Required = true;

            pd = schema.AddExpressionProperty("StartingColumn", "1");
            pd.DisplayName = "Starting Column";
            pd.Description = "Starting Column Property";
            pd.Required = true;

            // A repeat group of values to write out
            IRepeatGroupPropertyDefinition parts = schema.AddRepeatGroupProperty("Items");
            parts.Description = "The expression items to be written out.";

            pd = parts.PropertyDefinitions.AddExpressionProperty("Expression", String.Empty);
            pd.Description = "Expression value to be written out.";
        }

        /// <summary>
        /// Method called to create a new instance of this step type to place in a process. 
        /// Returns an instance of the class implementing the IStep interface.
        /// </summary>
        public IStep CreateStep(IPropertyReaders properties)
        {
            return new ExcelWriteStep(properties);
        }

        #endregion
    }

    class ExcelWriteStep : IStep
    {
        // As the step is created during the simulation run-time, get references
        // to the property readers that we'll need during the step Execution (for efficiency).
        IPropertyReaders _props;
        IPropertyReader _worksheetProp;
        IPropertyReader _rowProp;
        IPropertyReader _startingColumnProp;
        IElementProperty _ExcelconnectElementProp;
        IRepeatingPropertyReader _items;

        public ExcelWriteStep(IPropertyReaders properties)
        {
            _props = properties;
            _worksheetProp = _props.GetProperty("Worksheet");
            _rowProp = _props.GetProperty("Row");
            _startingColumnProp = _props.GetProperty("StartingColumn");
            _ExcelconnectElementProp = (IElementProperty)_props.GetProperty("ExcelConnectEPPlus");
            _items = (IRepeatingPropertyReader)_props.GetProperty("Items");
        }

        #region IStep Members

        /// <summary>
        /// Method called when a process token executes the step.
        /// </summary>
        public ExitType Execute(IStepExecutionContext context)
        {
            
            // Get an array of double values from the repeat group's list of expressions
            object[] paramsArray = new object[_items.GetCount(context)];
            for (int i = 0; i < _items.GetCount(context); i++)
            {
                // The thing returned from GetRow is IDisposable, so we use the using() pattern here
                using (IPropertyReaders row = _items.GetRow(i, context))
                {
                    // Get the expression property
                    IExpressionPropertyReader expressionProp = row.GetProperty("Expression") as IExpressionPropertyReader;
                    // Resolve the expression to get the value
                    paramsArray[i] = expressionProp.GetExpressionValue(context);
                }
            }           
            
            // set Excel data
            ExcelConnectElementEPPlus Excelconnect = (ExcelConnectElementEPPlus)_ExcelconnectElementProp.GetElement(context);
            if (Excelconnect == null)
            {
                context.ExecutionInformation.ReportError("ExcelConnectEPPlus element is null.  Makes sure ExcelWorkbook is defined correctly.");
            }
            String worksheetString = _worksheetProp.GetStringValue(context);
            Int32 rowInt = (Int32)_rowProp.GetDoubleValue(context);
            Int32 startingColumnInt = Convert.ToInt32(_startingColumnProp.GetDoubleValue(context));           

            try
            {
                // for each parameter
                for (int ii = 0; ii < paramsArray.Length; ii++)
                {
                    double doubleValue = TryAsDouble((Convert.ToString(paramsArray[ii],CultureInfo.InvariantCulture)));
                    if (! System.Double.IsNaN(doubleValue))
                    {
                        Excelconnect.WriteResults(worksheetString, rowInt, startingColumnInt + ii, doubleValue, DateTime.MinValue, String.Empty, context);                       
                    }
                    else
                    {
                        DateTime datetimeValue = TryAsDateTime((Convert.ToString(paramsArray[ii], CultureInfo.InvariantCulture)));
                        if (datetimeValue > System.DateTime.MinValue)
                        {
                            Excelconnect.WriteResults(worksheetString, rowInt, startingColumnInt + ii, System.Double.MinValue, datetimeValue, String.Empty, context);
                        }
                        else
                        {
                            Excelconnect.WriteResults(worksheetString, rowInt, startingColumnInt + ii, System.Double.MinValue, System.DateTime.MinValue, (Convert.ToString(paramsArray[ii], CultureInfo.InvariantCulture)), context);
                        }
                    }
                }
            }
            catch (FormatException)
            {
                context.ExecutionInformation.ReportError("Bad format provided in Excel Write step."); 
            }            

            // We are done writing, have the token proceed out of the primary exit
            return ExitType.FirstExit;
        }

        double TryAsDouble(string rawValue)
        {
            if (Double.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
            {
                return d;
            }
            else
            {
                return System.Double.NaN;
            }
        }

        DateTime TryAsDateTime(string rawValue)
        {
            if ( DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt))
            {
                return dt;                
            }
            else
            {
                return System.DateTime.MinValue; ;
            }
        }

        #endregion
    }
}
