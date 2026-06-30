using System;
using System.Windows.Forms;
using КП_БД_Черных.Forms;

namespace КП_БД_Черных
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new LoginForm());
        }
    }
}