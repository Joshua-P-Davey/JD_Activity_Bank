using System;
using System.IO;
using System.Threading; 
using System.Activities;
using System.ComponentModel;

namespace JD_Activity_Bank
{
    public class WaitFileExists : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("The full path to be checked, including the file name")]
        public InArgument<string> FilePath { get; set; }

        [Category("Input")]
        [Description("Specifies the amount of time (in milliseconds) to wait for " +
            "the activity to run before an error is thrown. The default value is 30000 " +
            "milliseconds (30 seconds).")]
        public InArgument<int> TimeoutMS { get; set; }

        [Category("Output")]
        [Description("States if the file was found within the allocated time.")]
        public OutArgument<bool> FileExists { get;  set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                //Assign default values
                bool out_boolFileExists = false;
                int intFrequency = 10;
                int intDefaultTimeout = 30000;
                int intWaitPeriod = 0;
                
                // Retrieves the necessary values from the properties pane 
                string in_StrFilePath = FilePath.Get(context);
                int in_intTimeoutMS = TimeoutMS.Get(context);

                // Check whether provided filepath value is whitespace 
                if (String.IsNullOrWhiteSpace(in_StrFilePath))
                {
                    throw (new Exception(String.Format("Provided Path Is Null Or WhiteSpace")));
                }

                // Check whether timeout value was provided and acquire absolute value 
                if (in_intTimeoutMS == 0)
                {
                    in_intTimeoutMS = intDefaultTimeout;
                }
                else
                {
                    in_intTimeoutMS = Math.Abs(in_intTimeoutMS);
                }

                // Setup the wait periods 
                intWaitPeriod = Convert.ToInt32(in_intTimeoutMS / intFrequency);

                // Setup timeout time
                DateTime dtTimeout = DateTime.Now.Add(TimeSpan.FromMilliseconds(in_intTimeoutMS));
                
                // As long as timeout time has not elapsed 
                while (DateTime.Now <= dtTimeout && !out_boolFileExists)
                {
                    if (File.Exists(in_StrFilePath))
                    {
                        out_boolFileExists = true;
                    }
                    else
                    {
                        Thread.Sleep(TimeSpan.FromMilliseconds(intWaitPeriod));
                    }
                }
                Console.WriteLine(out_boolFileExists);

                //Output resultant value
                FileExists.Set(context, out_boolFileExists);
            }
            catch (Exception exc)
            {
                throw (new System.Exception(string.Format("WaitFileExists Failed:- {0}Reason: {1} {0}Source: {2}", Environment.NewLine, exc.Message, exc.Source)));
            }
        }
    }
}
