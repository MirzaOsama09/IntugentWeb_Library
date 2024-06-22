using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.SqlServer.Server;
using IntugentClassLibrary.Pages.Mfg;
using System.Xml.Linq;
using System;
using IntugentWebApp.Utilities;
using IntugentClassLbrary.Classes;
using System.Drawing;
using System.Data.SqlClient;
using System.Data;
using System ;

namespace IntugentWebApp.Pages.Mfg_Group
{
    public class MfgProcessCheckModel : PageModel
    {

        public string gLoc1 { get; set; }
        public string gLoc2 {get;set;}
        public string gLoc3 {get;set;}
        public DateTime? gTestTime { get; set; }

        public readonly ObjectsService _objectsService;

        public MfgProcessCheckModel(ObjectsService objectsService)
        {
            _objectsService = objectsService;
        }
        public void OnGet()
        {
            int itmp;

            gLoc1  = _objectsService.MfgProcesscheck.CDefault.sLocMfg1;
            gLoc2  = _objectsService.MfgProcesscheck.CDefault.sLocMfg2;
            gLoc3  = _objectsService.MfgProcesscheck.CDefault.sLocMfg3;

            try
            {
                _objectsService.MfgProcesscheck.cbfile.conAZ.Open();
                sSqlQuery = "SELECT  top(1000) * FROM [dbo].[Process Check] where IDLocation = " + _objectsService.MfgProcesscheck.CDefault.IDLocation.ToString() + "  order by ID Desc  ";
                da = new SqlDataAdapter(sSqlQuery, _objectsService.MfgProcesscheck.cbfile.conAZ);

                if (_objectsService.MfgProcesscheck.dt == null) _objectsService.MfgProcesscheck.dt = new DataTable(); else _objectsService.MfgProcesscheck.dt.Clear();
                itmp = da.Fill(_objectsService.MfgProcesscheck.dt);
                if (itmp < 1)
                {
                 //   sMsgData = "There is no Process Check Data for " + _objectsService.MfgProcesscheck.CDefault.sLocation;
                    EnableDataControls(false);
                 //   CTelClient.TelTrace(sMsgData);
                    return;

                }
                drIndex = 0;
                dr = _objectsService.MfgProcesscheck.dt.Rows[0];
            }
            catch (Exception ex)
            {
               // sMsgData = "Error in contacting Process Check Data ";
               // System.Diagnostics.Trace.TraceError(sMsgData);
                EnableDataControls(false);
                gNewCheckSheet.IsEnabled = false;
              //  CTelClient.TelException(ex, sMsgData);  //Azue Insight Trace Message
                return;
            }
            finally
            {
                _objectsService.MfgProcesscheck.cbfile.conAZ.Close();
            }
            //  if (sMsgData != string.Empty) MessageBox.Show(sMsgData, _objectsService.MfgProcesscheck.cbfile.sAppName);
            SetView();
        }
        private void SetView()
        {
         
            string sTimeNullFormat = " ", sDateNullFormat = " "; string sTimeFormat = "hh:mm tt";

            if (_objectsService.MfgProcesscheck.drIndex == 0) gDataSetNext.IsEnabled = false; else gDataSetNext.IsEnabled = true;
            if (_objectsService.MfgProcesscheck.drIndex == _objectsService.MfgProcesscheck.dt.Rows.Count - 1) gDataSetPrev.IsEnabled = false; else gDataSetPrev.IsEnabled = true;

            double[] x = new double[] { 0.0, 1.0, 2.0, 3.0 };
            double[] y = new double[] { 0.0, 1.0, 2.0, 3.0 };
            //           gMark.Sources['X'].Data = x;
            //            gMark.Sources['Y'].Data = new double[] { 0.0, 1.0, 2.0, 3.0};
            //          var line1 = new InteractiveDataDisplay.WPF.MarkerGraph();

            //          gMark.Plot(x, y, x);


            if (_objectsService.MfgProcesscheck.dr == null) return;


            if (_objectsService.MfgProcesscheck.dr["ID"] == DBNull.Value) gID  = string.Empty; else gID  = _objectsService.MfgProcesscheck.dr["ID"].ToString();

            #region date, time, and datetime controls 

            if (_objectsService.MfgProcesscheck.dr["Sample Date Time"] == DBNull.Value) { gTestTime= DateTime.Today; gTestDate.SelectedDate = null; gTestDate  = "Select a Date"; }
            else { gTestTime = (DateTime)_objectsService.MfgProcesscheck.dr["Sample Date Time"]; gTestDate.SelectedDate = (DateTime)_objectsService.MfgProcesscheck.dr["Sample Date Time"]; }
            if ((_objectsService.MfgProcesscheck.dr["Product Code"] == DBNull.Value)) gProdID.SelectedValue = String.Empty; else gProdID.SelectedValue = _objectsService.MfgProcesscheck.dr["Product Code"].ToString();
            if ((_objectsService.MfgProcesscheck.dr["Operator"] == DBNull.Value)) gOperator.SelectedValue = -1; else gOperator.SelectedValue = (int)_objectsService.MfgProcesscheck.dr["Operator"];
            if (_objectsService.MfgProcesscheck.dr["Check Type"] == DBNull.Value) gType.SelectedValue = -1; else gType.SelectedValue = (int)_objectsService.MfgProcesscheck.dr["Check Type"];
            #endregion

            #region ComboBoxes
            if (_objectsService.MfgProcesscheck.dr["Product Code"] == DBNull.Value) gProdID.SelectedIndex = -1; else gProdID.SelectedValue = (string)_objectsService.MfgProcesscheck.dr["Product Code"];
            if (_objectsService.MfgProcesscheck.dr["Top Board Print"] == DBNull.Value) gTopBoardPrint.SelectedIndex = -1; else gTopBoardPrint.SelectedValue = (int)_objectsService.MfgProcesscheck.dr["Top Board Print"];
            if (_objectsService.MfgProcesscheck.dr["Bottom Board Print"] == DBNull.Value) gBottomBoardPrint.SelectedIndex = -1; else gBottomBoardPrint.SelectedValue = (int)_objectsService.MfgProcesscheck.dr["Bottom Board Print"];
            if (_objectsService.MfgProcesscheck.dr["Perferation"] == DBNull.Value) gPerferation.SelectedIndex = -1; else gPerferation.SelectedValue = (int)_objectsService.MfgProcesscheck.dr["Perferation"];
            if (_objectsService.MfgProcesscheck.dr["Flipper Operating"] == DBNull.Value) gFlipper.SelectedIndex = -1; else gFlipper.SelectedValue = (int)_objectsService.MfgProcesscheck.dr["Flipper Operating"];
            if (_objectsService.MfgProcesscheck.dr["Facer Adhesion"] == DBNull.Value) gAdhesion.SelectedIndex = -1; else gAdhesion.SelectedValue = (int)_objectsService.MfgProcesscheck.dr["Facer Adhesion"];
            if (_objectsService.MfgProcesscheck.dr["Edge Cut"] == DBNull.Value) gEdgeCut.SelectedIndex = -1; else gEdgeCut.SelectedValue = (int)_objectsService.MfgProcesscheck.dr["Edge Cut"];
            if (_objectsService.MfgProcesscheck.dr["Hooder Quality"] == DBNull.Value) gHooder.SelectedIndex = -1; else gHooder.SelectedValue = (int)_objectsService.MfgProcesscheck.dr["Hooder Quality"];
            if (_objectsService.MfgProcesscheck.dr["Board Quality"] == DBNull.Value) gBoardQuality.SelectedIndex = -1; else gBoardQuality.SelectedValue = (int)_objectsService.MfgProcesscheck.dr["Board Quality"];
            #endregion

            #region Bundle 1,2

            if (_objectsService.MfgProcesscheck.dr["Bundle Quantity 1"] == DBNull.Value) gQuantity  = string.Empty; else gQuantity  = _objectsService.MfgProcesscheck.dr["Bundle Quantity 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Bundle Wi_objectsService.MfgProcesscheck.dth 1"] == DBNull.Value) gWi_objectsService.MfgProcesscheck.dth  = string.Empty; else gWi_objectsService.MfgProcesscheck.dth  = _objectsService.MfgProcesscheck.dr["Bundle Wi_objectsService.MfgProcesscheck.dth 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Top Board Length 1"] == DBNull.Value) gTopLength  = string.Empty; else gTopLength  = _objectsService.MfgProcesscheck.dr["Top Board Length 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Middle Board Length 1"] == DBNull.Value) gMiddleLength  = string.Empty; else gMiddleLength  = _objectsService.MfgProcesscheck.dr["Middle Board Length 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Bottom Board Length 1"] == DBNull.Value) gBottomLength  = string.Empty; else gBottomLength  = _objectsService.MfgProcesscheck.dr["Bottom Board Length 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Diagonal_1 1"] == DBNull.Value) gDiagonal1  = string.Empty; else gDiagonal1  = _objectsService.MfgProcesscheck.dr["Diagonal_1 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Diagonal_2 1"] == DBNull.Value) gDiagonal2  = string.Empty; else gDiagonal2  = _objectsService.MfgProcesscheck.dr["Diagonal_2 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Wi_objectsService.MfgProcesscheck.dth Average 1"] == DBNull.Value) gWi_objectsService.MfgProcesscheck.dthAvg_1  = string.Empty; else gWi_objectsService.MfgProcesscheck.dthAvg_1  = ((double)_objectsService.MfgProcesscheck.dr["Wi_objectsService.MfgProcesscheck.dth Average 1"]).ToString(sFormat);
            if (_objectsService.MfgProcesscheck.dr["Squareness 1"] == DBNull.Value) gSquareness_1  = string.Empty; else gSquareness_1  = ((double)_objectsService.MfgProcesscheck.dr["Squareness 1"]).ToString(sFormat);
            if (_objectsService.MfgProcesscheck.dr["Length Average 1"] == DBNull.Value) gLengthAvg_1  = string.Empty; else gLengthAvg_1  = ((double)_objectsService.MfgProcesscheck.dr["Length Average 1"]).ToString(sFormat);

            if (_objectsService.MfgProcesscheck.dr["Bundle Quantity 2"] == DBNull.Value) gQuantity_2  = string.Empty; else gQuantity_2  = _objectsService.MfgProcesscheck.dr["Bundle Quantity 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Bundle Wi_objectsService.MfgProcesscheck.dth 2"] == DBNull.Value) gWi_objectsService.MfgProcesscheck.dth_2  = string.Empty; else gWi_objectsService.MfgProcesscheck.dth_2  = _objectsService.MfgProcesscheck.dr["Bundle Wi_objectsService.MfgProcesscheck.dth 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Top Board Length 2"] == DBNull.Value) gTopLength_2  = string.Empty; else gTopLength_2  = _objectsService.MfgProcesscheck.dr["Top Board Length 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Middle Board Length 2"] == DBNull.Value) gMiddleLength_2  = string.Empty; else gMiddleLength_2  = _objectsService.MfgProcesscheck.dr["Middle Board Length 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Bottom Board Length 2"] == DBNull.Value) gBottomLength_2  = string.Empty; else gBottomLength_2  = _objectsService.MfgProcesscheck.dr["Bottom Board Length 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Diagonal_1 2"] == DBNull.Value) gDiagonal1_2  = string.Empty; else gDiagonal1_2  = _objectsService.MfgProcesscheck.dr["Diagonal_1 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Diagonal_2 2"] == DBNull.Value) gDiagonal2_2  = string.Empty; else gDiagonal2_2  = _objectsService.MfgProcesscheck.dr["Diagonal_2 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Wi_objectsService.MfgProcesscheck.dth Average 2"] == DBNull.Value) gWi_objectsService.MfgProcesscheck.dthAvg_2  = string.Empty; else gWi_objectsService.MfgProcesscheck.dthAvg_2  = ((double)_objectsService.MfgProcesscheck.dr["Wi_objectsService.MfgProcesscheck.dth Average 2"]).ToString(sFormat);
            if (_objectsService.MfgProcesscheck.dr["Squareness 2"] == DBNull.Value) gSquareness_2  = string.Empty; else gSquareness_2  = ((double)_objectsService.MfgProcesscheck.dr["Squareness 2"]).ToString(sFormat);
            if (_objectsService.MfgProcesscheck.dr["Length Average 2"] == DBNull.Value) gLengthAvg_2  = string.Empty; else gLengthAvg_2  = ((double)_objectsService.MfgProcesscheck.dr["Length Average 2"]).ToString(sFormat);

            #endregion
            #region Board

            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc1"] == DBNull.Value) gThickness1  = string.Empty; else gThickness1  = _objectsService.MfgProcesscheck.dr["ThicknessLoc1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc2"] == DBNull.Value) gThickness2  = string.Empty; else gThickness2  = _objectsService.MfgProcesscheck.dr["ThicknessLoc2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc3"] == DBNull.Value) gThickness3  = string.Empty; else gThickness3  = _objectsService.MfgProcesscheck.dr["ThicknessLoc3"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Thickness Average"] == DBNull.Value) gThicknessAvg  = string.Empty; else gThicknessAvg  = ((double)_objectsService.MfgProcesscheck.dr["Thickness Average"]).ToString(sFormat);
            if (_objectsService.MfgProcesscheck.dr["Taper"] == DBNull.Value) gTaper  = string.Empty; else gTaper  = ((double)_objectsService.MfgProcesscheck.dr["Taper"]).ToString(sFormat);

            #endregion

            #region Misc Controls
            if (_objectsService.MfgProcesscheck.dr["bExclude"] == DBNull.Value) gExclude.IsChecked = false; else gExclude.IsChecked = (bool)_objectsService.MfgProcesscheck.dr["bExclude"];
            if (_objectsService.MfgProcesscheck.dr["Comment"] == DBNull.Value) gComment  = string.Empty; else gComment  = _objectsService.MfgProcesscheck.dr["Comment"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Exposed Foam"] == DBNull.Value) gExposedFoam  = string.Empty; else gExposedFoam  = _objectsService.MfgProcesscheck.dr["Exposed Foam"].ToString();
            #endregion

            #region Reactivity
            if (_objectsService.MfgProcesscheck.dr["EmptyCupMassG"] == DBNull.Value) gEmptyCupMassG  = string.Empty; else gEmptyCupMassG  = _objectsService.MfgProcesscheck.dr["EmptyCupMassG"].ToString();
            if (_objectsService.MfgProcesscheck.dr["CreamTimeS"] == DBNull.Value) gCreamTimeS  = string.Empty; else gCreamTimeS  = _objectsService.MfgProcesscheck.dr["CreamTimeS"].ToString();
            if (_objectsService.MfgProcesscheck.dr["GelTimeS"] == DBNull.Value) gGelTimeS  = string.Empty; else gGelTimeS  = _objectsService.MfgProcesscheck.dr["GelTimeS"].ToString();
            if (_objectsService.MfgProcesscheck.dr["TackFreeTimeS"] == DBNull.Value) gTackFreeTimeS  = string.Empty; else gTackFreeTimeS  = _objectsService.MfgProcesscheck.dr["TackFreeTimeS"].ToString();
            if (_objectsService.MfgProcesscheck.dr["FullCupMassG"] == DBNull.Value) gFullCupMassG  = string.Empty; else gFullCupMassG  = _objectsService.MfgProcesscheck.dr["FullCupMassG"].ToString();
            if (_objectsService.MfgProcesscheck.dr["FoamDensityPCF"] == DBNull.Value) gFoamDensityPCF  = string.Empty; else gFoamDensityPCF  = ((double)_objectsService.MfgProcesscheck.dr["FoamDensityPCF"]).ToString(sFormat);

            #endregion

            #region Board Deviation
            if (_objectsService.MfgProcesscheck.dr["DeviationFromTableRel"] == DBNull.Value)
                gBoardDeviationRel  = string.Empty;
            else gBoardDeviationRel  = ((double)_objectsService.MfgProcesscheck.dr["DeviationFromTableRel"]).ToString(sFormat);

            if (_objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"] == DBNull.Value) gDeviationAbs  = string.Empty; else gDeviationAbs  = _objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"].ToString();
            if (_objectsService.MfgProcesscheck.dr["DeviationType"] == DBNull.Value) gDeviationType.SelectedIndex = -1; else gDeviationType  = _objectsService.MfgProcesscheck.dr["DeviationType"].ToString();
            #endregion
            CheckLimits("All");
        }
        private void CheckLimits(string sF)
        {
            //Must be included in setview and   change products

            //if (sF == "All") CPro_objectsService.MfgProcesscheck.dtargets.GetProductTargets();


            //if (sF == "gThicknessAvg" || sF == "All")
            //{
            //    if (dr["Thickness Average"] == DBNull.Value) gThicknessAvg.Background = backColorCal;
            //    else if (CPro_objectsService.MfgProcesscheck.dtargets.ThicknessWithinLimits((double)dr["Thickness Average"]) == "N") gThicknessAvg.Background = backColorWarn; else gThicknessAvg.Background = backColorCal;
            //}

            //return;

            //if (sF == "gLengthAvg_1" || sF == "All")
            //{
            //    if (dr["Length Average 1"] == DBNull.Value) gLengthAvg_1.Background = backColorCal;
            //    else if (CPro_objectsService.MfgProcesscheck.dtargets.LengthWithinLimits((double)dr["Length Average 1"]) == "N") gLengthAvg_1.Background = backColorWarn; else gLengthAvg_1.Background = backColorCal;
            //}

            //if (sF == "gWi_objectsService.MfgProcesscheck.dthAvg_1" || sF == "All")
            //{
            //    if (dr["Wi_objectsService.MfgProcesscheck.dth Average 1"] == DBNull.Value) gWi_objectsService.MfgProcesscheck.dthAvg_1.Background = backColorCal;
            //    else if (CPro_objectsService.MfgProcesscheck.dtargets.Wi_objectsService.MfgProcesscheck.dthWithinLimits((double)dr["Wi_objectsService.MfgProcesscheck.dth Average 1"]) == "N") gWi_objectsService.MfgProcesscheck.dthAvg_1.Background = backColorWarn; else gWi_objectsService.MfgProcesscheck.dthAvg_1.Background = backColorCal;
            //}

            //if (sF == "gSquareness_1" || sF == "All")
            //{
            //    if (dr["Squareness 1"] == DBNull.Value) gSquareness_1.Background = backColorCal;
            //    else if (CPro_objectsService.MfgProcesscheck.dtargets.SquarenessWithinLimits((double)dr["Squareness 1"]) == "N") gSquareness_1.Background = backColorWarn; else gSquareness_1.Background = backColorCal;
            //}

            //if (sF == "gLengthAvg_2" || sF == "All")
            //{
            //    if (dr["Length Average 2"] == DBNull.Value) gLengthAvg_2.Background = backColorCal;
            //    else if (CPro_objectsService.MfgProcesscheck.dtargets.LengthWithinLimits((double)dr["Length Average 2"]) == "N") gLengthAvg_2.Background = backColorWarn; else gLengthAvg_2.Background = backColorCal;
            //}

            //if (sF == "gWi_objectsService.MfgProcesscheck.dthAvg_2" || sF == "All")
            //{
            //    if (dr["Wi_objectsService.MfgProcesscheck.dth Average 2"] == DBNull.Value) gWi_objectsService.MfgProcesscheck.dthAvg_2.Background = backColorCal;
            //    else if (CPro_objectsService.MfgProcesscheck.dtargets.Wi_objectsService.MfgProcesscheck.dthWithinLimits((double)dr["Wi_objectsService.MfgProcesscheck.dth Average 2"]) == "N") gWi_objectsService.MfgProcesscheck.dthAvg_2.Background = backColorWarn; else gWi_objectsService.MfgProcesscheck.dthAvg_2.Background = backColorCal;
            //}

            //if (sF == "gSquareness_2" || sF == "All")
            //{
            //    if (dr["Squareness 2"] == DBNull.Value) gSquareness_2.Background = backColorCal;
            //    else if (CPro_objectsService.MfgProcesscheck.dtargets.SquarenessWithinLimits((double)dr["Squareness 2"]) == "N") gSquareness_2.Background = backColorWarn; else gSquareness_2.Background = backColorCal;
            //}
        }
        private void gDeviation_LF(object sender, RoutedEventArgs e)
        {
            Control ctl = sender as Control;
            if (ctl == null) return;
            switch (ctl.Name)
            {
                case "gDeviationAbs": GetFloatField(gDeviationAbs, "DeviationFromTableAbs"); break;
                case "gDeviationType": if (gDeviationType.SelectedIndex > -1) _objectsService.MfgProcesscheck.dr["DeviationType"] = gDeviationType ; else _objectsService.MfgProcesscheck.dr["DeviationType"] = DBNull.Value; break;
            }

            if (_objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"] != DBNull.Value && _objectsService.MfgProcesscheck.dr["DeviationType"] != DBNull.Value)
            {
                if ((string)_objectsService.MfgProcesscheck.dr["DeviationType"] == "Up") _objectsService.MfgProcesscheck.dr["DeviationFromTableRel"] = _objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"];
                else _objectsService.MfgProcesscheck.dr["DeviationFromTableRel"] = -1.0 * (double)_objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"];
                gBoardDeviationRel  = ((double)_objectsService.MfgProcesscheck.dr["DeviationFromTableRel"]).ToString(sFormat);
            }
            else { _objectsService.MfgProcesscheck.dr["DeviationFromTableRel"] = DBNull.Value; gBoardDeviationRel  = string.Empty; }

            UpdateDataSet();
        }
        private void EnableDataControls(bool bstate = true)
        {
            gPrint.IsEnabled = bstate;
            gGenInfo.IsEnabled = bstate;
            gNavigation.IsEnabled = bstate;
            gFacer.IsEnabled = bstate;
            gBoard.IsEnabled = bstate;
            gBundle1.IsEnabled = bstate;
            gBundle2.IsEnabled = bstate;
            gBundleAvg.IsEnabled = bstate;
            gSPCComb.IsEnabled = bstate;
            gComment.IsEnabled = bstate;
            gCupReactivity.IsEnabled = bstate;
            gBoardDeviation.IsEnabled = bstate;
        }

        private void gNewCheckSheet_Click(object sender, RoutedEventArgs e)
        {
            string sMsg; int itmp;
            try
            {
                _objectsService.MfgProcesscheck.cbfile.conAZ.Open();
                string sql = "Select Next Value for [dbo].[IDProcessCheckSeq]";
                SqlCommand cmd = new SqlCommand(sql, _objectsService.MfgProcesscheck.cbfile.conAZ);
                itmp = (int)cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                sMsg = "Error in getting ID # of new process check datasheet";
              //  MessageBox.Show(sMsg, _objectsService.MfgProcesscheck.cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Stop);
                sMsg = "Could not create a new sequence number for RND Dataset";
                System.Diagnostics.Trace.TraceError(sMsg + "\n\n" + ex.Message);
                //CTelClient.TelException(ex, sMsg);
                return;
            }
            finally { _objectsService.MfgProcesscheck.cbfile.conAZ.Close(); }



            try
            {
                _objectsService.MfgProcesscheck.dr = _objectsService.MfgProcesscheck.dt.NewRow();                   //Create a new record
                _objectsService.MfgProcesscheck.dr["ID"] = itmp;
                _objectsService.MfgProcesscheck.dr["Operator"] = _objectsService.MfgProcesscheck.CDefault.IDEmployee;
                _objectsService.MfgProcesscheck.dr["IDLocation"] = _objectsService.MfgProcesscheck.CDefault.IDLocation;
                _objectsService.MfgProcesscheck.dr["Sample Date Time"] = DateTime.Now;// DBNull.Value; //
                _objectsService.MfgProcesscheck.dt.Rows.InsertAt(_objectsService.MfgProcesscheck.dr, 0);

                new SqlCommandBuilder(da);
                da.Update(_objectsService.MfgProcesscheck.dt);
                //                _objectsService.MfgProcesscheck.dt.DefaultView.Sort = "ID DESC";
                //                _objectsService.MfgProcesscheck.dt = _objectsService.MfgProcesscheck.dt.DefaultView.ToTable();
                _objectsService.MfgProcesscheck.dr = _objectsService.MfgProcesscheck.dt.Rows[0];
                EnableDataControls(true);
            }
            catch (Exception ex)
            {
               // sMsg = "Error in Creating new process check datasheet";
              //  MessageBox.Show(sMsg, _objectsService.MfgProcesscheck.cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Stop);
                sMsg = "Could not create a new Process Check Data with ID " + itmp.ToString();
                System.Diagnostics.Trace.TraceError(sMsg + "\n\n" + ex.Message);
               // CTelClient.TelException(ex, sMsg);
                EnableDataControls(false);
                return;
            }
            finally { _objectsService.MfgProcesscheck.cbfile.conAZ.Close(); }
           // sMsgData = string.Empty;

            SetView();

        } //Get seq. #

        private void gFirstBundle_LF(object sender, RoutedEventArgs e)
        {
            int ncount = 0; double dSum = 0.0;
            TextBox txtb = (TextBox)sender; bool bLength = false, bDiag = false, bWi_objectsService.MfgProcesscheck.dth = false;
            switch (txtb.Name)
            {
                case "gQuantity": GetIntField(txtb, "Bundle Quantity 1"); break;
                case "gWi_objectsService.MfgProcesscheck.dth": GetFloatField(txtb, "Bundle Wi_objectsService.MfgProcesscheck.dth 1"); bWi_objectsService.MfgProcesscheck.dth = true; break;
                case "gTopLength": GetFloatField(txtb, "Top Board Length 1"); bLength = true; break;
                case "gMiddleLength": GetFloatField(txtb, "Middle Board Length 1"); bLength = true; break;
                case "gBottomLength": GetFloatField(txtb, "Bottom Board Length 1"); bLength = true; break;
                case "gDiagonal1": GetFloatField(txtb, "Diagonal_1 1"); bDiag = true; break;
                case "gDiagonal2": GetFloatField(txtb, "Diagonal_2 1"); bDiag = true; break;


            }
            if (bWi_objectsService.MfgProcesscheck.dth)
            {
                if (_objectsService.MfgProcesscheck.dr["Bundle Wi_objectsService.MfgProcesscheck.dth 1"] != DBNull.Value) { gWi_objectsService.MfgProcesscheck.dthAvg_1  = ((double)_objectsService.MfgProcesscheck.dr["Bundle Wi_objectsService.MfgProcesscheck.dth 1"]).ToString(sFormat); _objectsService.MfgProcesscheck.dr["Wi_objectsService.MfgProcesscheck.dth Average 1"] = _objectsService.MfgProcesscheck.dr["Bundle Wi_objectsService.MfgProcesscheck.dth 1"]; }
                else { gWi_objectsService.MfgProcesscheck.dthAvg_1  = string.Empty; _objectsService.MfgProcesscheck.dr["Wi_objectsService.MfgProcesscheck.dth Average 1"] = DBNull.Value; }
                CheckLimits("gWi_objectsService.MfgProcesscheck.dthAvg_1");
            }
            else if (bDiag)
            {
                if (_objectsService.MfgProcesscheck.dr["Diagonal_1 1"] != DBNull.Value && _objectsService.MfgProcesscheck.dr["Diagonal_2 1"] != DBNull.Value)
                { _objectsService.MfgProcesscheck.dr["Squareness 1"] = Math.Abs((double)_objectsService.MfgProcesscheck.dr["Diagonal_1 1"] - (double)_objectsService.MfgProcesscheck.dr["Diagonal_2 1"]); gSquareness_1  = ((double)_objectsService.MfgProcesscheck.dr["Squareness 1"]).ToString(sFormat); }
                else { _objectsService.MfgProcesscheck.dr["Squareness 1"] = DBNull.Value; gSquareness_1  = String.Empty; }
                CheckLimits("gSquareness_1");

            }
            else if (bLength)
            {
                if (_objectsService.MfgProcesscheck.dr["Top Board Length 1"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Top Board Length 1"]; }
                if (_objectsService.MfgProcesscheck.dr["Middle Board Length 1"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Middle Board Length 1"]; }
                if (_objectsService.MfgProcesscheck.dr["Bottom Board Length 1"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Bottom Board Length 1"]; }
                if (ncount > 0) { dSum = dSum / (double)ncount; _objectsService.MfgProcesscheck.dr["Length Average 1"] = dSum; gLengthAvg_1  = dSum.ToString(sFormat); }
                else { _objectsService.MfgProcesscheck.dr["Length Average 1"] = DBNull.Value; gLengthAvg_1  = string.Empty; }
                CheckLimits("gLengthAvg_1");

            }
            UpdateDataSet();
        }

        private void gSecondBundle_LF(object sender, RoutedEventArgs e)
        {
            int ncount = 0; double dSum = 0.0;
            TextBox txtb = (TextBox)sender; bool bLength = false, bDiag = false, bWi_objectsService.MfgProcesscheck.dth = false;
            switch (txtb.Name)
            {
                case "gQuantity_2": GetIntField(txtb, "Bundle Quantity 2"); break;
                case "gWi_objectsService.MfgProcesscheck.dth_2": GetFloatField(txtb, "Bundle Wi_objectsService.MfgProcesscheck.dth 2"); bWi_objectsService.MfgProcesscheck.dth = true; break;
                case "gTopLength_2": GetFloatField(txtb, "Top Board Length 2"); bLength = true; break;
                case "gMiddleLength_2": GetFloatField(txtb, "Middle Board Length 2"); bLength = true; break;
                case "gBottomLength_2": GetFloatField(txtb, "Bottom Board Length 2"); bLength = true; break;
                case "gDiagonal1_2": GetFloatField(txtb, "Diagonal_1 2"); bDiag = true; break;
                case "gDiagonal2_2": GetFloatField(txtb, "Diagonal_2 2"); bDiag = true; break;


            }
            if (bWi_objectsService.MfgProcesscheck.dth)
            {
                if (_objectsService.MfgProcesscheck.dr["Bundle Wi_objectsService.MfgProcesscheck.dth 2"] != DBNull.Value) { gWi_objectsService.MfgProcesscheck.dthAvg_2  = ((double)_objectsService.MfgProcesscheck.dr["Bundle Wi_objectsService.MfgProcesscheck.dth 2"]).ToString(sFormat); _objectsService.MfgProcesscheck.dr["Wi_objectsService.MfgProcesscheck.dth Average 2"] = _objectsService.MfgProcesscheck.dr["Bundle Wi_objectsService.MfgProcesscheck.dth 2"]; }
                else { gWi_objectsService.MfgProcesscheck.dthAvg_2  = string.Empty; _objectsService.MfgProcesscheck.dr["Wi_objectsService.MfgProcesscheck.dth Average 2"] = DBNull.Value; }
                CheckLimits("gWi_objectsService.MfgProcesscheck.dthAvg_2");

            }
            else if (bDiag)
            {
                if (_objectsService.MfgProcesscheck.dr["Diagonal_1 2"] != DBNull.Value && _objectsService.MfgProcesscheck.dr["Diagonal_2 2"] != DBNull.Value)
                { _objectsService.MfgProcesscheck.dr["Squareness 2"] = Math.Abs((double)_objectsService.MfgProcesscheck.dr["Diagonal_1 2"] - (double)_objectsService.MfgProcesscheck.dr["Diagonal_2 2"]); gSquareness_2  = ((double)_objectsService.MfgProcesscheck.dr["Squareness 2"]).ToString(sFormat); }
                else { _objectsService.MfgProcesscheck.dr["Squareness 2"] = DBNull.Value; gSquareness_2  = String.Empty; }
                CheckLimits("gSquareness_2");

            }
            else if (bLength)
            {
                if (_objectsService.MfgProcesscheck.dr["Top Board Length 2"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Top Board Length 2"]; }
                if (_objectsService.MfgProcesscheck.dr["Middle Board Length 2"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Middle Board Length 2"]; }
                if (_objectsService.MfgProcesscheck.dr["Bottom Board Length 2"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Bottom Board Length 2"]; }
                if (ncount > 0) { dSum = dSum / (double)ncount; _objectsService.MfgProcesscheck.dr["Length Average 2"] = dSum; gLengthAvg_2  = dSum.ToString(sFormat); }
                else { _objectsService.MfgProcesscheck.dr["Length Average 2"] = DBNull.Value; gLengthAvg_2  = string.Empty; }
                CheckLimits("gLengthAvg_2");

            }
            UpdateDataSet();
        }
        private void gBoar_objectsService.MfgProcesscheck.dthickness_LF(object sender, RoutedEventArgs e)
        {
            int ncount = 0; double dSum = 0.0;
            TextBox txtb = (TextBox)sender; bool bLength = false, bDiag = false, bWi_objectsService.MfgProcesscheck.dth = false;
            switch (txtb.Name)
            {
                case "gThickness1": GetFloatField(txtb, "ThicknessLoc1"); break;
                case "gThickness2": GetFloatField(txtb, "ThicknessLoc2"); bWi_objectsService.MfgProcesscheck.dth = true; break;
                case "gThickness3": GetFloatField(txtb, "ThicknessLoc3"); bLength = true; break;
            }
            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc1"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["ThicknessLoc1"]; }
            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc2"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["ThicknessLoc2"]; }
            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc3"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["ThicknessLoc3"]; }
            if (ncount > 0) { dSum = dSum / (double)ncount; _objectsService.MfgProcesscheck.dr["Thickness Average"] = dSum; gThicknessAvg  = dSum.ToString(sFormat); }
            else { _objectsService.MfgProcesscheck.dr["Thickness Average"] = DBNull.Value; gThicknessAvg  = string.Empty; }

            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc1"] != DBNull.Value && _objectsService.MfgProcesscheck.dr["ThicknessLoc3"] != DBNull.Value)
            { dSum = Math.Abs((double)_objectsService.MfgProcesscheck.dr["ThicknessLoc3"] - (double)_objectsService.MfgProcesscheck.dr["ThicknessLoc1"]); _objectsService.MfgProcesscheck.dr["Taper"] = dSum; gTaper  = dSum.ToString(sFormat); }
            else { _objectsService.MfgProcesscheck.dr["Taper"] = DBNull.Value; gTaper  = string.Empty; }

        }



        private void gCupReactivity_LF(object sender, RoutedEventArgs e)
        {
            int ncount = 0; double dSum = 0.0;
            double _objectsService.MfgProcesscheck.dtmp1, _objectsService.MfgProcesscheck.dtmp2;
            TextBox txtb = (TextBox)sender; bool bMass = false;
            switch (txtb.Name)
            {
                case "gEmptyCupMassG": GetFloatField(txtb, "EmptyCupMassG"); bMass = true; break;
                case "gCreamTimeS": GetFloatField(txtb, "CreamTimeS"); break;
                case "gGelTimeS": GetFloatField(txtb, "GelTimeS"); break;
                case "gTackFreeTimeS": GetFloatField(txtb, "TackFreeTimeS"); break;
                case "gFullCupMassG": GetFloatField(txtb, "FullCupMassG"); bMass = true; break;
            }

            if (bMass)
            {
                if (_objectsService.MfgProcesscheck.dr["EmptyCupMassG"] != DBNull.Value && _objectsService.MfgProcesscheck.dr["FullCupMassG"] != DBNull.Value)
                {
                    _objectsService.MfgProcesscheck.dtmp1 = ((double)_objectsService.MfgProcesscheck.dr["FullCupMassG"] - (double)_objectsService.MfgProcesscheck.dr["EmptyCupMassG"]) / 453.592 / 32 * 957.506;
                    gFoamDensityPCF  = _objectsService.MfgProcesscheck.dtmp1.ToString(sFormat); _objectsService.MfgProcesscheck.dr["FoamDensityPCF"] = _objectsService.MfgProcesscheck.dtmp1;
                }
                else
                {
                    gFoamDensityPCF  = string.Empty; _objectsService.MfgProcesscheck.dr["FoamDensityPCF"] = DBNull.Value;
                }
            }

            UpdateDataSet();
        }
        private void gComboBox_LF(object sender, RoutedEventArgs e)
        {
            ComboBox cmb = sender as ComboBox;
            if (cmb == null) return;
            if (cmb.SelectedIndex < 0) return;

            switch (cmb.Name)
            {
                case "gProdID": _objectsService.MfgProcesscheck.dr["Product Code"] = cmb.SelectedValue; break;
                case "gOperator": _objectsService.MfgProcesscheck.dr["Operator"] = cmb.SelectedValue; break;
                case "gType": _objectsService.MfgProcesscheck.dr["Check Type"] = cmb.SelectedValue; break;
                case "gTopBoardPrint": _objectsService.MfgProcesscheck.dr["Top Board Print"] = cmb.SelectedValue; break;
                case "gBottomBoardPrint": _objectsService.MfgProcesscheck.dr["Bottom Board Print"] = cmb.SelectedValue; break;
                case "gPerferation": _objectsService.MfgProcesscheck.dr["Perferation"] = cmb.SelectedValue; break;
                case "gFlipper": _objectsService.MfgProcesscheck.dr["Flipper Operating"] = cmb.SelectedValue; break;
                case "gAdhesion": _objectsService.MfgProcesscheck.dr["Facer Adhesion"] = cmb.SelectedValue; break;
                case "gEdgeCut": _objectsService.MfgProcesscheck.dr["Edge Cut"] = cmb.SelectedValue; break;
                case "gHooder": _objectsService.MfgProcesscheck.dr["Hooder Quality"] = cmb.SelectedValue; break;
                case "gBoardQuality": _objectsService.MfgProcesscheck.dr["Board Quality"] = cmb.SelectedValue; break;
            }
            UpdateDataSet();
        }


        private void gMisc_LF(object sender, RoutedEventArgs e)
        {
            Control ctl = sender as Control;
            if (ctl == null) return;
            //           DateTime _objectsService.MfgProcesscheck.dt;

            switch (ctl.Name)
            {
                case "gTestDate":
                    if (gTestDate.SelectedDate != null)
                    {
                        if (_objectsService.MfgProcesscheck.dr["Sample Date Time"] == DBNull.Value) _objectsService.MfgProcesscheck.dr["Sample Date Time"] = gTestDate.SelectedDate.Value.Date;
                        else _objectsService.MfgProcesscheck.dr["Sample Date Time"] = gTestDate.SelectedDate.Value.Date + ((DateTime)_objectsService.MfgProcesscheck.dr["Sample Date Time"]).TimeOfDay;
                    }
                    break;
                case "gComment": _objectsService.MfgProcesscheck.dr["Comment"] = gComment ; break;
                case "gExclude": _objectsService.MfgProcesscheck.dr["bExclude"] = gExclude.IsChecked; break;
                case "gExposedFoam": GetFloatField(sender as TextBox, "Exposed Foam"); break;

            }
            UpdateDataSet();

        }
        private void gTestTime_LF(object sender, EventArgs e)
        {

            if (gTestTime != null)
            {
                if (_objectsService.MfgProcesscheck.dr["Sample Date Time"] == DBNull.Value) _objectsService.MfgProcesscheck.dr["Sample Date Time"] = ((DateTime)gTestTime);
                else _objectsService.MfgProcesscheck.dr["Sample Date Time"] = ((DateTime)_objectsService.MfgProcesscheck.dr["Sample Date Time"]).Date + ((DateTime)gTestTime).TimeOfDay;
            }
            UpdateDataSet();
        }
        private void CopyDataSet(object sender, RoutedEventArgs e)
        {
            string s_objectsService.MfgProcesscheck.dt1, s_objectsService.MfgProcesscheck.dt2; string sql;
            int itmp; DataTable _objectsService.MfgProcesscheck.dtCopy = null; SqlDataAdapter daCopy = null;

            if (gCopyData.SelectedIndex < 0) { MessageBox.Show("Choose an apprpriate time window", _objectsService.MfgProcesscheck.cbfile.sAppName); return; }
            DateTime _objectsService.MfgProcesscheck.dt1 = DateTime.Now;

            int isql = gCopyData.SelectedIndex;
            //            sql = "SELECT  * FROM [dbo].[Process Check] where IDLocation = " + _objectsService.MfgProcesscheck.CDefault.IDLocation.ToString();

            sql = "SELECT RN.ID, RN.[Product Code], RN.[Sample Date Time], R1.Employees, R2.sName as 'Check Type', R3.sName as 'Top Board Print', R4.sName as 'Bottom Board Print', R5.sName as 'Perferation', R6.sName as 'Flipper Operating', R7.sName as '[Facer Adhesion]', R8.sName as 'Edge Cut', R9.sName as 'Hooder Quality', R10.sName as 'Board Quality', R11.sName as 'Process Check Type', RN.Comment, RN.[Exposed Foam], RN.[Bundle Quantity 1] as 'Bundle 1 - Board Quantity', RN.[Bundle Wi_objectsService.MfgProcesscheck.dth 1] as 'Bundle 1 - Wi_objectsService.MfgProcesscheck.dth', RN.[Top Board Length 1] as 'Bundle 1 - Top Board Length', RN.[Middle Board Length 1] as 'Bundle 1 - Middle Board Length', RN.[Bottom Board Length 1] as 'Bundle 1 - Bottom Board Length', RN.[Diagonal_1 1] as 'Bundle 1 - Diagonal 1',RN.[Diagonal_2 1] as 'Bundle 1 - Diagonal 2', RN.[Length Average 1] as 'Bundle 1 - Average Length', RN.[Wi_objectsService.MfgProcesscheck.dth Average 1] as 'Bundle 1 - Average Wi_objectsService.MfgProcesscheck.dth', RN.[Squareness 1] as 'Bundle 1 - Squareness', RN.[Bundle Quantity 2] as 'Bundle 2 - Board Quantity', RN.[Bundle Wi_objectsService.MfgProcesscheck.dth 2] 'Bundle 2 - Wi_objectsService.MfgProcesscheck.dth', RN.[Top Board Length 2] as 'Bundle 2 - Top Board Length', RN.[Middle Board Length 2] as 'Bundle 2 - Middle Board Length', RN.[Bottom Board Length 2] as 'Bundle 2 - Bottom Board Length', RN.[Diagonal_1 2] as 'Bundle 2 - Diagonal 1', RN.[Diagonal_2 2] as 'Bundle 2 - Diagonal 2', RN.[Length Average 2] as 'Bundle 2 - Average Length', RN.[Wi_objectsService.MfgProcesscheck.dth Average 2] as 'Bundle 2 - Average Wi_objectsService.MfgProcesscheck.dth', RN.[Squareness 2] as 'Bundle 2 - Squareness', RN.ThicknessLoc1 as 'Board Thickness Location 1', RN.ThicknessLoc2 as 'Board Thickness Location 2', RN.ThicknessLoc3 as 'Board Thickness Location 3', RN.[Thickness Average] as 'Board Thickness Average', RN.Taper as 'Board Taper', R12.sLocation as 'Location', case when RN.bExclude = 1 then 'true' else 'false' end as 'Excluded from Anlysis if 1' " +
                "FROM[dbo].[Process Check] as RN Left Join[Roster] as R1 on RN.Operator = R1.ID Left Join tblLists as R2 on RN.[Check Type] = R2.ID Left Join tblLists as R3 on RN.[Top Board Print] = R3.ID Left Join tblLists as R4 on RN.[Bottom Board Print] = R4.ID Left Join tblLists as R5 on RN.Perferation = R5.ID Left Join tblLists as R6 on RN.[Flipper Operating] = R6.ID Left Join tblLists as R7 on RN.[Facer Adhesion] = R7.ID Left Join tblLists as R8 on RN.[Edge Cut] = R8.ID Left Join tblLists as R9 on RN.[Hooder Quality] = R9.ID Left Join tblLists as R10 on RN.[Board Quality] = R10.ID Left Join tblLists as R11 on RN.[Process Check Type] = R11.ID Left Join tblLocations as R12 on RN.IDLocation = R12.ID ";


            switch (isql)
            {
                case 0: _objectsService.MfgProcesscheck.dt1 = _objectsService.MfgProcesscheck.dt1.Date; sql += " And [Sample Date Time] >= '" + _objectsService.MfgProcesscheck.dt1.ToString() + "' And [Sample Date Time] < '" + _objectsService.MfgProcesscheck.dt1.AddDays(1).ToString() + "'"; break;
                case 1: sql += " And [Sample Date Time] <= '" + _objectsService.MfgProcesscheck.dt1.ToString() + "' And [Sample Date Time] > '" + _objectsService.MfgProcesscheck.dt1.AddDays(-1).ToString() + "'"; break;
                case 2: sql += " And [Sample Date Time] <= '" + _objectsService.MfgProcesscheck.dt1.ToString() + "' And [Sample Date Time] > '" + _objectsService.MfgProcesscheck.dt1.AddDays(-7).ToString() + "'"; break;
                case 3: sql += " And [Sample Date Time] <= '" + _objectsService.MfgProcesscheck.dt1.ToString() + "' And [Sample Date Time] > '" + _objectsService.MfgProcesscheck.dt1.AddMonths(-1).ToString() + "'"; break;
                case 4: sql += " And [Sample Date Time] <= '" + _objectsService.MfgProcesscheck.dt1.ToString() + "' And [Sample Date Time] > '" + _objectsService.MfgProcesscheck.dt1.AddMonths(-6).ToString() + "'"; break;
                case 5: sql += " And [Sample Date Time] <= '" + _objectsService.MfgProcesscheck.dt1.ToString() + "' And [Sample Date Time] > '" + _objectsService.MfgProcesscheck.dt1.AddYears(-1).ToString() + "'"; break;
                case 6: break;
                default: return;
            }
            sql += " and rn.idlocation = " + _objectsService.MfgProcesscheck.CDefault.IDLocation.ToString() + " order by [Sample Date Time] Desc  ";
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                _objectsService.MfgProcesscheck.cbfile.conAZ.Open();
                daCopy = new SqlDataAdapter(sql, _objectsService.MfgProcesscheck.cbfile.conAZ);
                if (_objectsService.MfgProcesscheck.dtCopy == null) _objectsService.MfgProcesscheck.dtCopy = new DataTable(); else _objectsService.MfgProcesscheck.dtCopy.Reset();
                itmp = da.Fill(_objectsService.MfgProcesscheck.dtCopy);
                if (itmp < 1)
                {
                    sMsgData = "There is no Process Check Data for you " + _objectsService.MfgProcesscheck.CDefault.sLocation + "during the selected time frame";
                    return;

                }
            }
            catch (Exception ex)
            {
                sMsgData = "Error in contacting Process Check Data for copying to clipboard ";
                MessageBox.Show(sMsgData, _objectsService.MfgProcesscheck.cbfile.sAppName);
                System.Diagnostics.Trace.TraceError(sMsgData);
                CTelClient.TelException(ex, sMsgData);  //Azue Insight Trace Message
                return;
            }
            finally
            {
                _objectsService.MfgProcesscheck.cbfile.conAZ.Close();
            }

            var sData = new StringBuilder();

            sData.Append(_objectsService.MfgProcesscheck.dtCopy.Columns[0].ColumnName.ToString());
            for (int icol = 1; icol < _objectsService.MfgProcesscheck.dtCopy.Columns.Count; icol++) sData.Append("\t" + _objectsService.MfgProcesscheck.dtCopy.Columns[icol].ColumnName.ToString());
            for (int irow = 0; irow < _objectsService.MfgProcesscheck.dtCopy.Rows.Count; irow++)
            {
                sData.Append("\n" + _objectsService.MfgProcesscheck.dtCopy.Rows[irow][0].ToString());
                for (int icol = 1; icol < _objectsService.MfgProcesscheck.dtCopy.Columns.Count; icol++) sData.Append("\t" + (_objectsService.MfgProcesscheck.dtCopy.Rows[irow][icol]).ToString());
            }
            Clipboard.SetText(sData.ToString());
            Mouse.OverrideCursor = null;
            CStatusBar.SetText("Search results copied to clipboard at " + DateTime.Now.ToString("hh:mm:ss:tt"));


        }
    }
}
