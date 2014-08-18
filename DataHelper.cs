using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ProgOpe
{
    public class DataHelper
    {
        public static void ErrorLog(string error)
        {
            try
            {
                //Pass the filepath and filename to the StreamWriter Constructor
                StreamWriter sw = new StreamWriter(Path.GetFullPath("./ProgErrorLogs.log"), true, Encoding.GetEncoding(932));

                //Write a line of text
                sw.WriteLine(string.Format("{0} : {1}", DateTime.Now, error));

                //Close the file
                sw.Close();
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
