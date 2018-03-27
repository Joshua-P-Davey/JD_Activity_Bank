using System;
using System.Activities;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace JD_Activity_Bank
{
    public class FormatRange : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> WorkbookFilePath { get; set; }

        [Category("Output")]
        public OutArgument<string[]> RetrievedSheets { get;  set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                // Retrieves the filepath value from the properties pane 
                string in_StrWorkbookFilePath = WorkbookFilePath.Get(context);

                // Check whether provided filepath value is appropriate 
                if (String.IsNullOrWhiteSpace(in_StrWorkbookFilePath) || !Uri.IsWellFormedUriString(in_StrWorkbookFilePath, UriKind.RelativeOrAbsolute))
                {
                    throw (new Exception(String.Format("Provided Path: {0}{1} {0}Is Either Null Or Not Valid, Please Review", Environment.NewLine, in_StrWorkbookFilePath)));
                }

                // Initialises a new instance of Excel 
                var xlInstance = new Application();

                // Open the provided workbook in readonly mode 
                xlInstance.Workbooks.Open(in_StrWorkbookFilePath, 0, true);

                // Initialise a string[] to store the retrieved sheetnames 
                string[] arrSheetNames = new string[xlInstance.Sheets.Count];

                // Initialise a counter to be used for Array index
                int intCounter = 0; 

                //For Each sheet in the workbook, add to string[]   
                foreach(Worksheet sheetName in xlInstance.Worksheets)
                {
                    arrSheetNames[intCounter] = sheetName.Name;
                    intCounter++;
                }

                // Close the provided workbook 
                xlInstance.Workbooks.Close();
            
                // Set the value of the output, provided in the properties 
                RetrievedSheets.Set(context, arrSheetNames);
            }
            catch (Exception exc)
            {
                throw (new System.Exception(string.Format("RetrieveWorkbookSheets Failed:- {0}Reason: {1} {0}Source: {2}", Environment.NewLine, exc.Message, exc.Source)));
            }
        }
    }
}
