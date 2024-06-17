using IntugentClassLbrary.Classes;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntugentClassLibrary.Utilities
{
    public class CLists_UpdateEmployee
    {
        public static void UpdateEmployee(CLists cLists)                     //Used to save user choices
        {
            //            SqlDataAdapter da = new SqlDataAdapter();

            SqlDataAdapter daEmployee = cLists.daEmployee;
            DataTable dtEmployee = cLists.dtEmployee;

            string sMsg = "Coult not save to the server";

            try
            {
                SqlCommandBuilder sb = new SqlCommandBuilder(daEmployee);
                sb.ConflictOption = ConflictOption.OverwriteChanges;
                daEmployee.Update(dtEmployee);
            }
            catch (Exception ex)
            {
                sMsg = "Errors in saving User preferences";
                //MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Stop);
                System.Diagnostics.Trace.TraceError(sMsg + "\n\n" + ex.Message);
                //CTelClient.TelException(ex, sMsg);
            }
            //CStatusBar.SetText("User Preferences Saved at " + DateTime.Now.ToString("hh:mm:ss:tt"));

        }
    }
}
