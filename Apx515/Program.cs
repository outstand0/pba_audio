#define BLOCK_DUPE_EXECUTE

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Apx515
{
    static class Program
    {
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
#if (!BLOCK_DUPE_EXECUTE)
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
#else
            #region block_dupe_exec
            bool createdNew;

            System.Threading.Mutex mutex = new System.Threading.Mutex(true, Application.ProductName, out createdNew);
            if(createdNew)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            else
            {
                MessageBox.Show("NG - This program is already running");
            }
            #endregion
#endif
        }
    }
}
