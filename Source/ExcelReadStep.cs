using System;
using System.Globalization;
using SimioAPI;
using SimioAPI.Extensions;

namespace ExcelReadWriteEPPlus
{
    class ExcelReadStepDefinition : IStepDefinition
    {
        #region IStepDefinition Members

        /// <summary>
        /// Property returning the full name for this type of step. The name should contain no spaces. 
        /// </summary>
        public string Name
        {
            get { return "ExcelReadEPPlus"; }
        }

        /// <summary>
        /// Property returning a short description of what the step does.  
        /// </summary>
        public string Description
        {
            get { return "The Excel Read step may be used to read values from an Excel worksheet using EPPlus."; }
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
        static readonly Guid MY_ID = new Guid("{758BD4CE-1884-404D-B96C-6E5C46640900}"); // Feb2020/dth

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

            // Reference to the file to read from
            pd = schema.AddElementProperty("ExcelConnectEPPlus", ExcelConnectEPPlusElementDefinition.MY_ID);

            pd = schema.AddStringProperty("Worksheet", String.Empty);
            pd.Description = "Worksheet Property";
            pd.Required = true;

            pd = schema.AddExpressionProperty("Row", "1");
            pd.Description = "Row Property";
            pd.Required = true;

            pd = schema.AddExpressionProperty("StartingColumn", "1");
            pd.DisplayName = "Starting Column";
            pd.Description = "Column Property";
            pd.Required = true;

            // A repeat group of states to read into
            IRepeatGroupPropertyDefinition parts = schema.AddRepeatGroupProperty("States");
            parts.Description = "The state values to read the values into";

            pd = parts.PropertyDefinitions.AddStateProperty("State");
            pd.Description = "A state to read a value into from Excel.";
        }

        /// <summary>
        /// Method called to create a new instance of this step type to place in a process. 
        /// Returns an instance of the class implementing the IStep interface.
        /// </summary>
        public IStep CreateStep(IPropertyReaders properties)
        {
            return new ExcelReadStep(properties);
        }

        #endregion
    }

    class ExcelReadStep : IStep
    {
        IPropertyReaders _props;
        IPropertyReader _worksheetProp;
        IPropertyReader _rowProp;
        IPropertyReader _startingColumnProp;
        IElementProperty _ExcelconnectElementProp;
        IRepeatingPropertyReader _states;
        public ExcelReadStep(IPropertyReaders properties)
        {
            _props = properties;
            _worksheetProp = _props.GetProperty("Worksheet");
            _rowProp = _props.GetProperty("Row");
            _startingColumnProp = _props.GetProperty("StartingColumn");
            _ExcelconnectElementProp = (IElementProperty)_props.GetProperty("ExcelConnect");
            _states = (IRepeatingPropertyReader)_props.GetProperty("States");
        }

        #region IStep Members

        /// <summary>
        /// Method called when a process token executes the step.
        /// </summary>
        public ExitType Execute(IStepExecutionContext context)
        {
            // Get Excel data
            ExcelConnectElementEPPlus Excelconnect = (ExcelConnectElementEPPlus)_ExcelconnectElementProp.GetElement(context);
            if (Excelconnect == null)
            {
                context.ExecutionInformation.ReportError("ExcelConnectEPPlus element is null.  Makes sure ExcelWorkbook is defined correctly.");
            }
            String worksheetString = _worksheetProp.GetStringValue(context);
            Int32 rowInt = (Int32)_rowProp.GetDoubleValue(context);
            Int32 startingColumnInt = Convert.ToInt32(_startingColumnProp.GetDoubleValue(context));           

            int numReadIn = 0;
            int numReadFailures = 0;
            for (int i = 0; i < _states.GetCount(context); i++)
            {
                // Tokenize the input
                string resultsString = Excelconnect.ReadResults(worksheetString, rowInt, startingColumnInt + i, context);

                // The thing returned from GetRow is IDisposable, so we use the using() pattern here
                using (IPropertyReaders row = _states.GetRow(i, context))
                {
                    // Get the state property out of the i-th tuple of the repeat group
                    IStateProperty stateprop = (IStateProperty)row.GetProperty("State");
                    // Resolve the property value to get the runtime state
                    IState state = stateprop.GetState(context);

                    if (TryAsNumericState(state, resultsString) ||
                        TryAsDateTimeState(state, resultsString) ||
                        TryAsStringState(state, resultsString))
                    {
                        numReadIn++;
                    }
                    else
                    {
                        numReadFailures++;
                    }
                }
            }

            string worksheetName = (_worksheetProp as IPropertyReader).GetStringValue(context);
            context.ExecutionInformation.TraceInformation($"Read from row={rowInt} worksheet={worksheetName} into {numReadIn} state columns. {numReadFailures} read failures");

            // We are done reading, have the token proceed out of the primary exit
            return ExitType.FirstExit;
        }

        bool TryAsNumericState(IState state, string rawValue)
        {
            IRealState realState = state as IRealState;
            if (realState == null)
                return false; // destination state is not a real.

            double d = 0.0;
            if (Double.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
            {
                realState.Value = d;
                return true;
            }
            else if (String.Compare(rawValue, "True", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                realState.Value = 1.0;
                return true;
            }
            else if (String.Compare(rawValue, "False", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                realState.Value = 0.0;
                return true;
            }

            return false; // incoming value can't be interpreted as a real.
        }

        bool TryAsDateTimeState(IState state, string rawValue)
        {
            IDateTimeState dateTimeState = state as IDateTimeState;
            if (dateTimeState == null)
                return false; // destination state is not a DateTime.

            DateTime dt;
            if (DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
            {
                dateTimeState.Value = dt;
                return true;
            }

            // If it isn't a DateTime, maybe it is just a number, which we can interpret as hours from start of simulation.
            double d = 0.0;
            if (Double.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
            {
                state.StateValue = d;
                return true;
            }

            return false;
        }

        bool TryAsStringState(IState state, string rawValue)
        {
            IStringState stringState = state as IStringState;
            if (stringState == null)
                return false; // destination state is not a string.

            // Since all input value are already strings, this is easy.
            stringState.Value = rawValue;
            return true;
        }

        #endregion
    }
}
