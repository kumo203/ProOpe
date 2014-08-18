using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ProgOpe
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                //二重起動をチェックする
                if (System.Diagnostics.Process.GetProcessesByName(System.Diagnostics.Process.GetCurrentProcess().ProcessName).Length > 1)
                {
                    DataHelper.ErrorLog(string.Format("Program Already Running"));
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new ProgOpeForm());
            }
            catch (Exception ex)
            {
                // Log Error
                DataHelper.ErrorLog(string.Format("{0} : {1}", DateTime.Now, ex.Message));
            }            
        }
    }
}
