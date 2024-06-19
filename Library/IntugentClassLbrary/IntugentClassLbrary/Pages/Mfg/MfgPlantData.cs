using IntugentClassLbrary.Classes;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntugentClassLibrary.Pages.Mfg
{
    public class MfgPlantData
    {
        string sSqlQuery;
        SqlDataAdapter da, da2;
       public DataTable dt = new DataTable(), dtPP = new DataTable();
        public DataRow dr, drIP, drFG;
        public DataTable dtPPChemDel = new DataTable(), dtPPChemDel1 = new DataTable(), dtPPPTable = new DataTable(), dtPPDBelt = new DataTable(), dtPPOthers = new DataTable(), dtNewInsData = new DataTable();
        public bool bDataSetChanged = false;
        public DateTime dateTime1, dateTime2, dtIPTime, dtFGTime, dtQCCheckTime;

        public Cbfile cBfile;
        public CLists clist;
        public MfgPlantData(Cbfile CBfile, CLists clist)
        {
            this.clist = clist;
            cBfile = CBfile;
            string sql = "SELECT * FROM [tblProcessParams]";
            da = new SqlDataAdapter(sql, cBfile.conAZ);
            int itmp = da.Fill(clist.dtProcessParams);

            clist.dvProcssParams = clist.dtProcessParams.DefaultView;
            clist.dvProcssParams.RowFilter = "sGroup = 'Chemical Delivery'";
            clist.dvPPChemDel = clist.dvProcssParams.ToTable().DefaultView;
            clist.dvProcssParams.RowFilter = "sGroup = 'Chemical Delivery 1'";
            clist.dvPPChemDel1 = clist.dvProcssParams.ToTable().DefaultView;

            clist.dvProcssParams.RowFilter = "sGroup = 'Pour Table'";
            clist.dvPPPTable = clist.dvProcssParams.ToTable().DefaultView;
            clist.dvProcssParams.RowFilter = "sGroup = 'Double Belt'";
            clist.dvPPDBelt = clist.dvProcssParams.ToTable().DefaultView;
            clist.dvProcssParams.RowFilter = "sGroup = 'New Instrument data - temp'";
            clist.dvNewInsData = clist.dvProcssParams.ToTable().DefaultView;

            clist.dvProcssParams.RowFilter = "sGroup NOT IN  ('Pour Table', 'Double Belt','Chemical Delivery', 'Chemical Delivery 1','New Instrument data - temp') ";
            clist.dvPPOthers = clist.dvProcssParams.ToTable().DefaultView;
        }


        public bool GetDataSet()
        {
            string sMsg, sn;

            if (dtPPChemDel != null) dtPPChemDel.Clear(); dtPPChemDel = clist.dvPPChemDel.ToTable();
            if (dtPPChemDel1 != null) dtPPChemDel1.Clear(); dtPPChemDel1 = clist.dvPPChemDel1.ToTable();
            if (dtPPPTable != null) dtPPPTable.Clear(); dtPPPTable = clist.dvPPPTable.ToTable();
            if (dtPPPTable != null) dtPPDBelt.Clear(); dtPPDBelt = clist.dvPPDBelt.ToTable();
            if (dtPPOthers != null) dtPPOthers.Clear(); dtPPOthers = clist.dvPPOthers.ToTable();
            if (dtNewInsData != null) dtNewInsData.Clear(); dtNewInsData = clist.dvNewInsData.ToTable();

            try
            {
                sSqlQuery = "Select * from [Process Data] where [ID4All]=" + cBfile.iIDMfg.ToString(); //1943";  //3137
                da = new SqlDataAdapter(sSqlQuery, cBfile.conAZ);

                dt.Clear();
                int itmp = da.Fill(dt);
                if (itmp < 1)
                {
                    sMsg = "Could not find the Process Data for the Selected Dataset";
                    //  MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Stop);
                    System.Diagnostics.Trace.TraceError(sMsg);
                    //  CPages.PageMfgHome_1.MfgDataNotFound();
                    return false;

                }
                dr = dt.Rows[0];
                for (int ir = 0; ir < 2; ir++)
                {
                    sn = ((string)dtPPChemDel.Rows[ir]["sName"]);
                    if (dr[sn] == DBNull.Value) dtPPChemDel.Rows[ir]["sValue"] = string.Empty; else dtPPChemDel.Rows[ir]["sValue"] = dr[sn].ToString();
                }

                for (int ir = 2; ir < dtPPChemDel.Rows.Count; ir++)
                {
                    sn = ((string)dtPPChemDel.Rows[ir]["sName"]);
                    if (dr[sn] == DBNull.Value) dtPPChemDel.Rows[ir]["sValue"] = string.Empty; else dtPPChemDel.Rows[ir]["sValue"] = ((double)dr[sn]).ToString("0.000");
                }

                for (int ir = 0; ir < dtPPChemDel1.Rows.Count; ir++)
                {
                    sn = ((string)dtPPChemDel1.Rows[ir]["sName"]);
                    if (dr[sn] == DBNull.Value) dtPPChemDel1.Rows[ir]["sValue"] = string.Empty; else dtPPChemDel1.Rows[ir]["sValue"] = ((double)dr[sn]).ToString("0.000");
                }

                for (int ir = 0; ir < dtPPPTable.Rows.Count; ir++)
                {
                    sn = ((string)dtPPPTable.Rows[ir]["sName"]);
                    if (dr[sn] == DBNull.Value) dtPPPTable.Rows[ir]["sValue"] = string.Empty; else dtPPPTable.Rows[ir]["sValue"] = ((double)dr[sn]).ToString("0.000");
                }

                for (int ir = 0; ir < dtPPDBelt.Rows.Count; ir++)
                {
                    sn = ((string)dtPPDBelt.Rows[ir]["sName"]);
                    if (dr[sn] == DBNull.Value) dtPPDBelt.Rows[ir]["sValue"] = string.Empty; else dtPPDBelt.Rows[ir]["sValue"] = ((double)dr[sn]).ToString("0.000");
                }

                for (int ir = 0; ir < dtPPOthers.Rows.Count; ir++)
                {
                    sn = ((string)dtPPOthers.Rows[ir]["sName"]);
                    if (dr[sn] == DBNull.Value) dtPPOthers.Rows[ir]["sValue"] = string.Empty; else dtPPOthers.Rows[ir]["sValue"] = ((double)dr[sn]).ToString("0.000");
                }

                for (int ir = 0; ir < dtNewInsData.Rows.Count; ir++)
                {
                    sn = ((string)dtNewInsData.Rows[ir]["sName"]);
                    if (dr[sn] == DBNull.Value) dtNewInsData.Rows[ir]["sValue"] = string.Empty; else dtNewInsData.Rows[ir]["sValue"] = ((double)dr[sn]).ToString("0.000");
                }

                //                CPages.PageInProcess_1.GetDataSet();
                //               drIP = CPages.PageInProcess_1.dr;
                //                CPages.PageFinishedGoods_1.GetDataSet();
                //               drFG = CPages.PageFinishedGoods_1.dr;

            }
            catch (SqlException ex)
            {
                sMsg = "Error in retrieving the Plant Data for the Selected Dataset";
                //MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Stop);
                System.Diagnostics.Trace.TraceError(sMsg + "\n\n" + ex.Message);
                // CPages.PageMfgHome_1.MfgDataNotFound();
                // CTelClient.TelException(ex, sMsg);
                return false;
            }



            return true;
        }

        public void UpdateDataSet()
        {
            string sMsg = "Coult not save to the server";
            try
            {
                SqlCommandBuilder sb = new SqlCommandBuilder(da);
                sb.ConflictOption = ConflictOption.OverwriteChanges;
                int v = da.Update(dt);
            }
            catch (Exception ex)
            {
                //  MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Stop);
                //sMsg = "Could not save the the process dataset " + Cbfile.iIDMfg.ToString();
                System.Diagnostics.Trace.TraceError(sMsg + "\n\n" + ex.Message);
                //  CTelClient.TelException(ex, sMsg);
                return;
            }

            //            CStatusBar.SetText("Data Saved at " + DateTime.Now.ToString("hh:mm:ss tt"));
        }
    }
 }