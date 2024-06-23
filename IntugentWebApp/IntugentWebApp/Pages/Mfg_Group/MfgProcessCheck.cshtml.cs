using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.SqlServer.Server;
using IntugentClassLibrary.Pages.Mfg;
using System.Xml.Linq;
using System;
using System.Text;
using IntugentWebApp.Utilities;
using IntugentClassLbrary.Classes;
using System.Drawing;
using System.Data.SqlClient;
using System.Data;
using Google.Api.Gax;

namespace IntugentWebApp.Pages.Mfg_Group
{
    [BindProperties]
    public class MfgProcessCheckModel : PageModel
    {
        public string gID { get; set; }
        public bool gDataSetNextIsEnabled { get; set; }
        public bool gDataSetPrevIsEnabled { get; set; }
        public string gQuantity { get; set; }
        public string gWidth { get; set; }
        public string gTopLength { get; set; }
        public string gMiddleLength { get; set; }
        public string gBottomLength { get; set; }
        public string gDiagonal1 { get; set; }
        public string gDiagonal2 { get; set; }
        public string gQuantity_2 { get; set; }
        public string gWidth_2 { get; set; }
        public string gTopLength_2 { get; set; }
        public string gMiddleLength_2 { get; set; }
        public string gBottomLength_2 { get; set; }
        public string gDiagonal1_2 { get; set; }
        public string gDiagonal2_2 { get; set; }
        public string gLengthAvg_1 { get; set; }
        public string gSquareness_1 { get; set; }
        public string gLengthAvg_2 { get; set; }
        public string gSquareness_2 { get; set; }
        public string gLoc1 { get; set; } 
        public string gLoc2 { get; set; }
        public string gLoc3 { get; set; }
        public string gThickness1 { get; set; }
        public string gThickness2 { get; set; }
        public string gThickness3 { get; set; }
        public string gThicknessAvg { get; set; }
        public string gTaper { get; set; }
        public string gWidthAvg_1 { get; set; }
        public string gWidthAvg_2 { get; set; }
        public string gEmptyCupMassG {get;set;}
        public string gCreamTimeS {get;set;}
        public string gGelTimeS {get;set;}
        public string gTackFreeTimeS {get;set;}
        public string gFullCupMassG {get;set;}
        public string gFoamDensityPCF{get;set;}
        public string gBoardDeviationRel { get; set; }
        public string gDeviationAbs { get; set; }
        public string gExposedFoam { get; set; }
        public string gComment { get; set; }
        public DateTime? gTestTime { get; set; }
        public DateTime? gTestDate {  get; set; }
        public DataView gProdID { get; set; }
        public string gProdIDSelected { get; set; }
        public DataView gType { get; set; }
        public int gTypeSelectedValue { get; set; }

        public DataView gOperator { get; set; }
        public int gOperatorSelectedItem { get; set; }
        public int gTopBoardPrint { get; set; }
        public DataView gTopBoardPrintCollection { get; set; }
        public int gBottomBoardPrint { get; set; }
        public DataView gBottomBoardPrintCol { get; set; }
        public int gPerferation { get; set; }
        public DataView gPerferationCol { get; set; }
        public int gFlipper { get; set; }
        public DataView gFlipperCol { get; set; }
        public int gAdhesionSelected { get; set; }
        public DataView gAdhesionCol { get; set; }
        public int gEdgeCutSelected { get; set; }
        public DataView gEdgeCutCol { get; set; }
        public int gHooderSelected { get; set; }
        public DataView gHooderCol { get; set; }
        public int gBoardQualitySelected { get; set; }
        public DataView gBoardQualityCol { get; set; }
        public string gDeviationType { get; set; }
        public int gCopyData {  get; set; }
        public bool gExcludeIsChecked {  get; set; }
        public bool gNewCheckSheetIsEnabled {  get; set; }
        public bool gPrintIsEnabled { get; set; }
        public bool gGenInfoIsEnabled { get; set; }
        public bool gNavigationIsEnabled {  get; set; }
        public bool gFacerIsEnabled {  get; set; }
        public bool gBoardIsEnabled {  get; set; }
        public bool gBundle1IsEnabled {  get; set; }
        public bool gBundle2IsEnabled {  get; set; }
        public bool gBundleAvgIsEnabled {  get; set; }
        public bool gSPCCombIsEnabled {  get; set; }
        public bool gCommentIsEnabled {  get; set; }
        public bool gCupReactivityIsEnabled {  get; set; }
        public bool gBoardDeviationIsEnabled {  get; set; }

        public readonly ObjectsService _objectsService;

        public MfgProcessCheckModel(ObjectsService objectsService)
        {
            _objectsService = objectsService;
        }
        public void OnGet()
        {
            gLoc1  = _objectsService.MfgProcesscheck.CDefault.sLocMfg1;
            gLoc2  = _objectsService.MfgProcesscheck.CDefault.sLocMfg2;
            gLoc3  = _objectsService.MfgProcesscheck.CDefault.sLocMfg3;

            int itmp;
            try
            {
                _objectsService.MfgProcesscheck.cbfile.conAZ.Open();
                _objectsService.MfgProcesscheck.sSqlQuery = "SELECT  top(1000) * FROM [dbo].[Process Check] where IDLocation = " + _objectsService.MfgProcesscheck.CDefault.IDLocation.ToString() + "  order by ID Desc  ";
                _objectsService.MfgProcesscheck.da = new SqlDataAdapter(_objectsService.MfgProcesscheck.sSqlQuery, _objectsService.MfgProcesscheck.cbfile.conAZ);

                if (_objectsService.MfgProcesscheck.dt == null) _objectsService.MfgProcesscheck.dt = new DataTable(); else _objectsService.MfgProcesscheck.dt.Clear();
                itmp = _objectsService.MfgProcesscheck.da.Fill(_objectsService.MfgProcesscheck.dt);
                if (itmp < 1)
                {
                 //   sMsgData = "There is no Process Check Data for " + _objectsService.MfgProcesscheck.CDefault.sLocation;
                    EnableDataControls(false);
                 //   CTelClient.TelTrace(sMsgData);
                    return;

                }
                _objectsService.MfgProcesscheck.drIndex = 0;
                _objectsService.MfgProcesscheck.dr = _objectsService.MfgProcesscheck.dt.Rows[0];
            }
            catch (Exception ex)
            {
               // sMsgData = "Error in contacting Process Check Data ";
               // System.Diagnostics.Trace.TraceError(sMsgData);
                EnableDataControls(false);
                gNewCheckSheetIsEnabled = false;
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
            gProdID = _objectsService.CLists.dvComProd;
            gProdIDSelected = _objectsService.CDefualts.sProdMfg;

            gType = _objectsService.CLists.dvRunType2;
            gTypeSelectedValue = _objectsService.CDefualts.iRunType;

            gOperator = _objectsService.CLists.dvEmployeesMfg;
            gOperatorSelectedItem = _objectsService.CDefualts.IDEmployee;

            gTopBoardPrintCollection = _objectsService.CLists.dvPCTopBoard;
            gTopBoardPrint = 0;

            gBottomBoardPrintCol = _objectsService.CLists.dvPCBottomBoard;
            gBottomBoardPrint = 0;

            gPerferationCol = _objectsService.CLists.dvPCPerferation;
            gPerferation = 0;

            gFlipperCol = _objectsService.CLists.dvPCFlipper;
            gFlipper = 0;

            gAdhesionCol= _objectsService.CLists.dvPCFacerAdh;
            gAdhesionSelected = 0;
            
            gEdgeCutCol =_objectsService.CLists.dvPCEdgeCut;
            gEdgeCutSelected = 0;
            
            gHooderCol= _objectsService.CLists.dvPCHooderQuality;
            gHooderSelected = 0;
            
            gBoardQualityCol = _objectsService.CLists.dvPCBoardQuality;
            gBoardQualitySelected = 0;

            string sTimeNullFormat = " ", sDateNullFormat = " "; string sTimeFormat = "hh:mm tt";

            if (_objectsService.MfgProcesscheck.drIndex == 0) gDataSetNextIsEnabled = false; else gDataSetNextIsEnabled = true;
            if (_objectsService.MfgProcesscheck.drIndex == _objectsService.MfgProcesscheck.dt.Rows.Count - 1) gDataSetPrevIsEnabled = false; else gDataSetPrevIsEnabled = true;

            double[] x = new double[] { 0.0, 1.0, 2.0, 3.0 };
            double[] y = new double[] { 0.0, 1.0, 2.0, 3.0 };
            //           gMark.Sources['X'].Data = x;
            //            gMark.Sources['Y'].Data = new double[] { 0.0, 1.0, 2.0, 3.0};
            //          var line1 = new InteractiveDataDisplay.WPF.MarkerGraph();

            //          gMark.Plot(x, y, x);


            if (_objectsService.MfgProcesscheck.dr == null) return;


            if (_objectsService.MfgProcesscheck.dr["ID"] == DBNull.Value) gID  = string.Empty; else gID  = _objectsService.MfgProcesscheck.dr["ID"].ToString();

            #region date, time, and datetime controls 

            if (_objectsService.MfgProcesscheck.dr["Sample Date Time"] == DBNull.Value) { gTestTime= DateTime.Today; gTestDate= null; gTestDate  = null; }
            else { gTestTime = (DateTime)_objectsService.MfgProcesscheck.dr["Sample Date Time"]; gTestDate = (DateTime)_objectsService.MfgProcesscheck.dr["Sample Date Time"]; }
            if ((_objectsService.MfgProcesscheck.dr["Product Code"] == DBNull.Value)) gProdIDSelected = String.Empty; else gProdIDSelected = _objectsService.MfgProcesscheck.dr["Product Code"].ToString();
            if ((_objectsService.MfgProcesscheck.dr["Operator"] == DBNull.Value)) gOperatorSelectedItem = -1; else gOperatorSelectedItem = (int)_objectsService.MfgProcesscheck.dr["Operator"];
            if (_objectsService.MfgProcesscheck.dr["Check Type"] == DBNull.Value) gTypeSelectedValue = -1; else gTypeSelectedValue = (int)_objectsService.MfgProcesscheck.dr["Check Type"];
            #endregion

            #region ComboBoxes
            if (_objectsService.MfgProcesscheck.dr["Product Code"] == DBNull.Value) gProdIDSelected = null; else gProdIDSelected = (string)_objectsService.MfgProcesscheck.dr["Product Code"];
            if (_objectsService.MfgProcesscheck.dr["Top Board Print"] == DBNull.Value) gTopBoardPrint = -1; else gTopBoardPrint= (int)_objectsService.MfgProcesscheck.dr["Top Board Print"];
            if (_objectsService.MfgProcesscheck.dr["Bottom Board Print"] == DBNull.Value) gBottomBoardPrint = -1; else gBottomBoardPrint = (int)_objectsService.MfgProcesscheck.dr["Bottom Board Print"];
            if (_objectsService.MfgProcesscheck.dr["Perferation"] == DBNull.Value) gPerferation = -1; else gPerferation = (int)_objectsService.MfgProcesscheck.dr["Perferation"];
            if (_objectsService.MfgProcesscheck.dr["Flipper Operating"] == DBNull.Value) gFlipper = -1; else gFlipper = (int)_objectsService.MfgProcesscheck.dr["Flipper Operating"];
            if (_objectsService.MfgProcesscheck.dr["Facer Adhesion"] == DBNull.Value) gAdhesionSelected = -1; else gAdhesionSelected = (int)_objectsService.MfgProcesscheck.dr["Facer Adhesion"];
            if (_objectsService.MfgProcesscheck.dr["Edge Cut"] == DBNull.Value) gEdgeCutSelected = -1; else gEdgeCutSelected = (int)_objectsService.MfgProcesscheck.dr["Edge Cut"];
            if (_objectsService.MfgProcesscheck.dr["Hooder Quality"] == DBNull.Value) gHooderSelected = -1; else gHooderSelected = (int)_objectsService.MfgProcesscheck.dr["Hooder Quality"];
            if (_objectsService.MfgProcesscheck.dr["Board Quality"] == DBNull.Value) gBoardQualitySelected = -1; else gBoardQualitySelected = (int)_objectsService.MfgProcesscheck.dr["Board Quality"];
            #endregion

            #region Bundle 1,2

            if (_objectsService.MfgProcesscheck.dr["Bundle Quantity 1"] == DBNull.Value) gQuantity  = string.Empty; else gQuantity  = _objectsService.MfgProcesscheck.dr["Bundle Quantity 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Bundle Width 1"] == DBNull.Value) gWidth  = string.Empty; else gWidth  = _objectsService.MfgProcesscheck.dr["Bundle Width 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Top Board Length 1"] == DBNull.Value) gTopLength  = string.Empty; else gTopLength  = _objectsService.MfgProcesscheck.dr["Top Board Length 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Middle Board Length 1"] == DBNull.Value) gMiddleLength  = string.Empty; else gMiddleLength  = _objectsService.MfgProcesscheck.dr["Middle Board Length 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Bottom Board Length 1"] == DBNull.Value) gBottomLength  = string.Empty; else gBottomLength  = _objectsService.MfgProcesscheck.dr["Bottom Board Length 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Diagonal_1 1"] == DBNull.Value) gDiagonal1  = string.Empty; else gDiagonal1  = _objectsService.MfgProcesscheck.dr["Diagonal_1 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Diagonal_2 1"] == DBNull.Value) gDiagonal2  = string.Empty; else gDiagonal2  = _objectsService.MfgProcesscheck.dr["Diagonal_2 1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Width Average 1"] == DBNull.Value) gWidthAvg_1  = string.Empty; else gWidthAvg_1  = ((double)_objectsService.MfgProcesscheck.dr["Width Average 1"]).ToString(_objectsService.MfgProcesscheck.sFormat);
            if (_objectsService.MfgProcesscheck.dr["Squareness 1"] == DBNull.Value) gSquareness_1  = string.Empty; else gSquareness_1  = ((double)_objectsService.MfgProcesscheck.dr["Squareness 1"]).ToString(_objectsService.MfgProcesscheck.sFormat);
            if (_objectsService.MfgProcesscheck.dr["Length Average 1"] == DBNull.Value) gLengthAvg_1  = string.Empty; else gLengthAvg_1  = ((double)_objectsService.MfgProcesscheck.dr["Length Average 1"]).ToString(_objectsService.MfgProcesscheck.sFormat);

            if (_objectsService.MfgProcesscheck.dr["Bundle Quantity 2"] == DBNull.Value) gQuantity_2  = string.Empty; else gQuantity_2  = _objectsService.MfgProcesscheck.dr["Bundle Quantity 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Bundle Width 2"] == DBNull.Value) gWidth_2  = string.Empty; else gWidth_2  = _objectsService.MfgProcesscheck.dr["Bundle Width 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Top Board Length 2"] == DBNull.Value) gTopLength_2  = string.Empty; else gTopLength_2  = _objectsService.MfgProcesscheck.dr["Top Board Length 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Middle Board Length 2"] == DBNull.Value) gMiddleLength_2  = string.Empty; else gMiddleLength_2  = _objectsService.MfgProcesscheck.dr["Middle Board Length 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Bottom Board Length 2"] == DBNull.Value) gBottomLength_2  = string.Empty; else gBottomLength_2  = _objectsService.MfgProcesscheck.dr["Bottom Board Length 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Diagonal_1 2"] == DBNull.Value) gDiagonal1_2  = string.Empty; else gDiagonal1_2  = _objectsService.MfgProcesscheck.dr["Diagonal_1 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Diagonal_2 2"] == DBNull.Value) gDiagonal2_2  = string.Empty; else gDiagonal2_2  = _objectsService.MfgProcesscheck.dr["Diagonal_2 2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Width Average 2"] == DBNull.Value) gWidthAvg_2  = string.Empty; else gWidthAvg_2  = ((double)_objectsService.MfgProcesscheck.dr["Width Average 2"]).ToString(_objectsService.MfgProcesscheck.sFormat);
            if (_objectsService.MfgProcesscheck.dr["Squareness 2"] == DBNull.Value) gSquareness_2  = string.Empty; else gSquareness_2  = ((double)_objectsService.MfgProcesscheck.dr["Squareness 2"]).ToString(_objectsService.MfgProcesscheck.sFormat);
            if (_objectsService.MfgProcesscheck.dr["Length Average 2"] == DBNull.Value) gLengthAvg_2  = string.Empty; else gLengthAvg_2  = ((double)_objectsService.MfgProcesscheck.dr["Length Average 2"]).ToString(_objectsService.MfgProcesscheck.sFormat);

            #endregion
            #region Board

            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc1"] == DBNull.Value) gThickness1  = string.Empty; else gThickness1  = _objectsService.MfgProcesscheck.dr["ThicknessLoc1"].ToString();
            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc2"] == DBNull.Value) gThickness2  = string.Empty; else gThickness2  = _objectsService.MfgProcesscheck.dr["ThicknessLoc2"].ToString();
            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc3"] == DBNull.Value) gThickness3  = string.Empty; else gThickness3  = _objectsService.MfgProcesscheck.dr["ThicknessLoc3"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Thickness Average"] == DBNull.Value) gThicknessAvg  = string.Empty; else gThicknessAvg  = ((double)_objectsService.MfgProcesscheck.dr["Thickness Average"]).ToString(_objectsService.MfgProcesscheck.sFormat);
            if (_objectsService.MfgProcesscheck.dr["Taper"] == DBNull.Value) gTaper  = string.Empty; else gTaper  = ((double)_objectsService.MfgProcesscheck.dr["Taper"]).ToString(_objectsService.MfgProcesscheck.sFormat);

            #endregion

            #region Misc Controls
            if (_objectsService.MfgProcesscheck.dr["bExclude"] == DBNull.Value) gExcludeIsChecked = false; else gExcludeIsChecked = (bool)_objectsService.MfgProcesscheck.dr["bExclude"];
            if (_objectsService.MfgProcesscheck.dr["Comment"] == DBNull.Value) gComment  = string.Empty; else gComment  = _objectsService.MfgProcesscheck.dr["Comment"].ToString();
            if (_objectsService.MfgProcesscheck.dr["Exposed Foam"] == DBNull.Value) gExposedFoam  = string.Empty; else gExposedFoam  = _objectsService.MfgProcesscheck.dr["Exposed Foam"].ToString();
            #endregion

            #region Reactivity
            if (_objectsService.MfgProcesscheck.dr["EmptyCupMassG"] == DBNull.Value) gEmptyCupMassG  = string.Empty; else gEmptyCupMassG  = _objectsService.MfgProcesscheck.dr["EmptyCupMassG"].ToString();
            if (_objectsService.MfgProcesscheck.dr["CreamTimeS"] == DBNull.Value)    gCreamTimeS  = string.Empty; else gCreamTimeS  = _objectsService.MfgProcesscheck.dr["CreamTimeS"].ToString();
            if (_objectsService.MfgProcesscheck.dr["GelTimeS"] == DBNull.Value)      gGelTimeS  = string.Empty; else gGelTimeS  = _objectsService.MfgProcesscheck.dr["GelTimeS"].ToString();
            if (_objectsService.MfgProcesscheck.dr["TackFreeTimeS"] == DBNull.Value) gTackFreeTimeS  = string.Empty; else gTackFreeTimeS  = _objectsService.MfgProcesscheck.dr["TackFreeTimeS"].ToString();
            if (_objectsService.MfgProcesscheck.dr["FullCupMassG"] == DBNull.Value) gFullCupMassG  = string.Empty; else gFullCupMassG  = _objectsService.MfgProcesscheck.dr["FullCupMassG"].ToString();
            if (_objectsService.MfgProcesscheck.dr["FoamDensityPCF"] == DBNull.Value)gFoamDensityPCF  = string.Empty; else gFoamDensityPCF  = ((double)_objectsService.MfgProcesscheck.dr["FoamDensityPCF"]).ToString(_objectsService.MfgProcesscheck.sFormat);

            #endregion

            #region Board Deviation
            if (_objectsService.MfgProcesscheck.dr["DeviationFromTableRel"] == DBNull.Value)
                gBoardDeviationRel  = string.Empty;
            else gBoardDeviationRel  = ((double)_objectsService.MfgProcesscheck.dr["DeviationFromTableRel"]).ToString(_objectsService.MfgProcesscheck.sFormat);

            if (_objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"] == DBNull.Value) gDeviationAbs  = string.Empty; else gDeviationAbs  = _objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"].ToString();
            if (_objectsService.MfgProcesscheck.dr["DeviationType"] == DBNull.Value) gDeviationType = null; else gDeviationType  = _objectsService.MfgProcesscheck.dr["DeviationType"].ToString();
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
        public IActionResult OnPostDeviation_LF(string Name,string value)
        {
            //Control ctl = sender as Control;
            //if (ctl == null) return;
            switch (Name)
            {
                case "gDeviationAbs": _objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"] = value; break;
                case "gDeviationType": 
                    if (value!=null) _objectsService.MfgProcesscheck.dr["DeviationType"] = value ; 
                    else _objectsService.MfgProcesscheck.dr["DeviationType"] = DBNull.Value; break;
            }

            if (_objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"] != DBNull.Value && _objectsService.MfgProcesscheck.dr["DeviationType"] != DBNull.Value)
            {
                if ((string)_objectsService.MfgProcesscheck.dr["DeviationType"] == "Up") _objectsService.MfgProcesscheck.dr["DeviationFromTableRel"] = _objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"];
                else _objectsService.MfgProcesscheck.dr["DeviationFromTableRel"] = -1.0 * (double)_objectsService.MfgProcesscheck.dr["DeviationFromTableAbs"];
                gBoardDeviationRel  = ((double)_objectsService.MfgProcesscheck.dr["DeviationFromTableRel"]).ToString(_objectsService.MfgProcesscheck.sFormat);
            }
            else { _objectsService.MfgProcesscheck.dr["DeviationFromTableRel"] = DBNull.Value; gBoardDeviationRel  = string.Empty; }

            _objectsService.MfgProcesscheck.UpdateDataSet();
            return new JsonResult(new { message = Name, value });
        }
        private void EnableDataControls(bool bstate = true)
        {
            gPrintIsEnabled = bstate;
            gGenInfoIsEnabled = bstate;
            gNavigationIsEnabled = bstate;
            gFacerIsEnabled = bstate;
            gBoardIsEnabled = bstate;
            gBundle1IsEnabled = bstate;
            gBundle2IsEnabled = bstate;
            gBundleAvgIsEnabled = bstate;
            gSPCCombIsEnabled = bstate;
            gCommentIsEnabled = bstate;
            gCupReactivityIsEnabled = bstate;
            gBoardDeviationIsEnabled = bstate;
        }

        public IActionResult OnPostGNewCheckSheet_Click()
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
                return null;
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

                new SqlCommandBuilder(_objectsService.MfgProcesscheck.da);
                _objectsService.MfgProcesscheck.da.Update(_objectsService.MfgProcesscheck.dt);
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
                return null;
            }
            finally { _objectsService.MfgProcesscheck.cbfile.conAZ.Close(); }
           // sMsgData = string.Empty;

            SetView();
            return new JsonResult(new {message= true});
        } //Get seq. #

        public IActionResult OnPostGFirstBundle_LF(string Name,string value)
        {
            int ncount = 0; double dSum = 0.0;
            bool bLength = false, bDiag = false, bWidth = false;
            switch (Name)
            {
                case "gQuantity": _objectsService.MfgProcesscheck.dr["Bundle Quantity 1"]=Int32.Parse(value); break;
                case "gWidth": _objectsService.MfgProcesscheck.dr["Bundle Width 1"]=value; bWidth = true; break;
                case "gTopLength": _objectsService.MfgProcesscheck.dr["Top Board Length 1"]=value; bLength = true; break;
                case "gMiddleLength":_objectsService.MfgProcesscheck.dr["Middle Board Length 1"]=value; bLength = true; break;
                case "gBottomLength":_objectsService.MfgProcesscheck.dr["Bottom Board Length 1"]=value; bLength = true; break;
                case "gDiagonal1": _objectsService.MfgProcesscheck.dr["Diagonal_1 1"] = value; bDiag = true; break;
                case "gDiagonal2": _objectsService.MfgProcesscheck.dr["Diagonal_2 1"]=value; bDiag = true; break;


            }
            if (bWidth)
            {
                if (_objectsService.MfgProcesscheck.dr["Bundle Width 1"] != DBNull.Value) { gWidthAvg_1  = ((double)_objectsService.MfgProcesscheck.dr["Bundle Width 1"]).ToString(_objectsService.MfgProcesscheck.sFormat); _objectsService.MfgProcesscheck.dr["Width Average 1"] = _objectsService.MfgProcesscheck.dr["Bundle Width 1"]; }
                else { gWidthAvg_1  = string.Empty; _objectsService.MfgProcesscheck.dr["Width Average 1"] = DBNull.Value; }
                CheckLimits("gWidthAvg_1");
            }
            else if (bDiag)
            {
                if (_objectsService.MfgProcesscheck.dr["Diagonal_1 1"] != DBNull.Value && _objectsService.MfgProcesscheck.dr["Diagonal_2 1"] != DBNull.Value)
                { _objectsService.MfgProcesscheck.dr["Squareness 1"] = Math.Abs((double)_objectsService.MfgProcesscheck.dr["Diagonal_1 1"] - (double)_objectsService.MfgProcesscheck.dr["Diagonal_2 1"]); gSquareness_1  = ((double)_objectsService.MfgProcesscheck.dr["Squareness 1"]).ToString(_objectsService.MfgProcesscheck.sFormat); }
                else { _objectsService.MfgProcesscheck.dr["Squareness 1"] = DBNull.Value; gSquareness_1  = String.Empty; }
                CheckLimits("gSquareness_1");

            }
            else if (bLength)
            {
                if (_objectsService.MfgProcesscheck.dr["Top Board Length 1"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Top Board Length 1"]; }
                if (_objectsService.MfgProcesscheck.dr["Middle Board Length 1"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Middle Board Length 1"]; }
                if (_objectsService.MfgProcesscheck.dr["Bottom Board Length 1"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Bottom Board Length 1"]; }
                if (ncount > 0) { dSum = dSum / (double)ncount; _objectsService.MfgProcesscheck.dr["Length Average 1"] = dSum; gLengthAvg_1  = dSum.ToString(_objectsService.MfgProcesscheck.sFormat); }
                else { _objectsService.MfgProcesscheck.dr["Length Average 1"] = DBNull.Value; gLengthAvg_1  = string.Empty; }
                CheckLimits("gLengthAvg_1");

            }
            _objectsService.MfgProcesscheck.UpdateDataSet();
            return new JsonResult(new { message = Name, value });
        }

        public IActionResult OnPostGSecondBundle_LF(string Name, string value)
        {
            int ncount = 0; double dSum = 0.0;
            bool bLength = false, bDiag = false, bWidth = false;
            switch (Name)
            {
                case "gQuantity_2": _objectsService.MfgProcesscheck.dr["Bundle Quantity 2"]=Int32.Parse(value); break;
                case "gWidth_2": _objectsService.MfgProcesscheck.dr["Bundle Width 2"] = value; bWidth = true; break;
                case "gTopLength_2": _objectsService.MfgProcesscheck.dr["Top Board Length 2"]=value; bLength = true; break;
                case "gMiddleLength_2": _objectsService.MfgProcesscheck.dr["Middle Board Length 2"]=value; bLength = true; break;
                case "gBottomLength_2": _objectsService.MfgProcesscheck.dr["Bottom Board Length 2"]=value; bLength = true; break;
                case "gDiagonal1_2": _objectsService.MfgProcesscheck.dr["Diagonal_1 2"]=value; bDiag = true; break;
                case "gDiagonal2_2": _objectsService.MfgProcesscheck.dr["Diagonal_2 2"]=value; bDiag = true; break;


            }
            if (bWidth)
            {
                if (_objectsService.MfgProcesscheck.dr["Bundle Width 2"] != DBNull.Value) { gWidthAvg_2  = ((double)_objectsService.MfgProcesscheck.dr["Bundle Width 2"]).ToString(_objectsService.MfgProcesscheck.sFormat); _objectsService.MfgProcesscheck.dr["Width Average 2"] = _objectsService.MfgProcesscheck.dr["Bundle Width 2"]; }
                else { gWidthAvg_2  = string.Empty; _objectsService.MfgProcesscheck.dr["Width Average 2"] = DBNull.Value; }
                CheckLimits("gWidthAvg_2");

            }
            else if (bDiag)
            {
                if (_objectsService.MfgProcesscheck.dr["Diagonal_1 2"] != DBNull.Value && _objectsService.MfgProcesscheck.dr["Diagonal_2 2"] != DBNull.Value)
                { _objectsService.MfgProcesscheck.dr["Squareness 2"] = Math.Abs((double)_objectsService.MfgProcesscheck.dr["Diagonal_1 2"] - (double)_objectsService.MfgProcesscheck.dr["Diagonal_2 2"]); gSquareness_2  = ((double)_objectsService.MfgProcesscheck.dr["Squareness 2"]).ToString(_objectsService.MfgProcesscheck.sFormat); }
                else { _objectsService.MfgProcesscheck.dr["Squareness 2"] = DBNull.Value; gSquareness_2  = String.Empty; }
                CheckLimits("gSquareness_2");

            }
            else if (bLength)
            {
                if (_objectsService.MfgProcesscheck.dr["Top Board Length 2"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Top Board Length 2"]; }
                if (_objectsService.MfgProcesscheck.dr["Middle Board Length 2"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Middle Board Length 2"]; }
                if (_objectsService.MfgProcesscheck.dr["Bottom Board Length 2"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["Bottom Board Length 2"]; }
                if (ncount > 0) { dSum = dSum / (double)ncount; _objectsService.MfgProcesscheck.dr["Length Average 2"] = dSum; gLengthAvg_2  = dSum.ToString(_objectsService.MfgProcesscheck.sFormat); }
                else { _objectsService.MfgProcesscheck.dr["Length Average 2"] = DBNull.Value; gLengthAvg_2  = string.Empty; }
                CheckLimits("gLengthAvg_2");

            }
            _objectsService.MfgProcesscheck.UpdateDataSet();
            return new JsonResult(new { message = Name, value });
        }
        public IActionResult OnPostBoardthickness_LF(string Name, string value)
        {
            int ncount = 0; double dSum = 0.0;
            bool bLength = false, bDiag = false, bWidth = false;
            switch (Name)
            {
                case "gThickness1": _objectsService.MfgProcesscheck.dr["ThicknessLoc1"]=value; break;
                case "gThickness2": _objectsService.MfgProcesscheck.dr["ThicknessLoc2"]=value; bWidth = true; break;
                case "gThickness3": _objectsService.MfgProcesscheck.dr["ThicknessLoc3"]=value; bLength = true; break;
            }
            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc1"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["ThicknessLoc1"]; }
            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc2"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["ThicknessLoc2"]; }
            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc3"] != DBNull.Value) { ncount += 1; dSum += (double)_objectsService.MfgProcesscheck.dr["ThicknessLoc3"]; }
            if (ncount > 0) { dSum = dSum / (double)ncount; _objectsService.MfgProcesscheck.dr["Thickness Average"] = dSum; gThicknessAvg  = dSum.ToString(_objectsService.MfgProcesscheck.sFormat); }
            else { _objectsService.MfgProcesscheck.dr["Thickness Average"] = DBNull.Value; gThicknessAvg  = string.Empty; }

            if (_objectsService.MfgProcesscheck.dr["ThicknessLoc1"] != DBNull.Value && _objectsService.MfgProcesscheck.dr["ThicknessLoc3"] != DBNull.Value)
            { dSum = Math.Abs((double)_objectsService.MfgProcesscheck.dr["ThicknessLoc3"] - (double)_objectsService.MfgProcesscheck.dr["ThicknessLoc1"]); _objectsService.MfgProcesscheck.dr["Taper"] = dSum; gTaper  = dSum.ToString(_objectsService.MfgProcesscheck.sFormat); }
            else { _objectsService.MfgProcesscheck.dr["Taper"] = DBNull.Value; gTaper  = string.Empty; }
            _objectsService.MfgProcesscheck.UpdateDataSet();
            return new JsonResult(new { message = Name, value });
        }



        public IActionResult OnPostGCupReactivity_LF(string Name, string value)
        {
            int ncount = 0; double dSum = 0.0;
            double dtmp1, dtmp2;
            bool bMass = false;
            switch (Name)
            {
                case "gEmptyCupMassG": _objectsService.MfgProcesscheck.dr["EmptyCupMassG"]=value; bMass = true; break;
                case "gCreamTimeS":    _objectsService.MfgProcesscheck.dr["CreamTimeS"] = value; break;
                case "gGelTimeS":      _objectsService.MfgProcesscheck.dr["GelTimeS"] = value; break;
                case "gTackFreeTimeS": _objectsService.MfgProcesscheck.dr["TackFreeTimeS"] = value; break;
                case "gFullCupMassG":  _objectsService.MfgProcesscheck.dr["FullCupMassG"] = value; bMass = true; break;
            }

            if (bMass)
            {
                if (_objectsService.MfgProcesscheck.dr["EmptyCupMassG"] != DBNull.Value && _objectsService.MfgProcesscheck.dr["FullCupMassG"] != DBNull.Value)
                {
                    dtmp1 = ((double)_objectsService.MfgProcesscheck.dr["FullCupMassG"] - (double)_objectsService.MfgProcesscheck.dr["EmptyCupMassG"]) / 453.592 / 32 * 957.506;
                    gFoamDensityPCF  = dtmp1.ToString(_objectsService.MfgProcesscheck.sFormat); _objectsService.MfgProcesscheck.dr["FoamDensityPCF"] = dtmp1;
                }
                else
                {
                    gFoamDensityPCF  = string.Empty; _objectsService.MfgProcesscheck.dr["FoamDensityPCF"] = DBNull.Value;
                }
            }

            _objectsService.MfgProcesscheck.UpdateDataSet();
            return new JsonResult(new { message = Name, value });
        }
        public IActionResult OnPostGComboBox_LF(string Name, string value)
        {
            //ComboBox cmb = sender as ComboBox;
            //if (cmb == null) return null;
            //if (cmb.SelectedIndex < 0) return null;

            switch (Name)
            {
                case "gProdID": _objectsService.MfgProcesscheck.dr["Product Code"] = value; break;
                case "gOperator": _objectsService.MfgProcesscheck.dr["Operator"] = value; break;
                case "gType": _objectsService.MfgProcesscheck.dr["Check Type"] = value; break;
                case "gTopBoardPrint": _objectsService.MfgProcesscheck.dr["Top Board Print"] = value; break;
                case "gBottomBoardPrint": _objectsService.MfgProcesscheck.dr["Bottom Board Print"] = value; break;
                case "gPerferation": _objectsService.MfgProcesscheck.dr["Perferation"] = value; break;
                case "gFlipper": _objectsService.MfgProcesscheck.dr["Flipper Operating"] = value; break;
                case "gAdhesion": _objectsService.MfgProcesscheck.dr["Facer Adhesion"] = value; break;
                case "gEdgeCut": _objectsService.MfgProcesscheck.dr["Edge Cut"] = value; break;
                case "gHooder": _objectsService.MfgProcesscheck.dr["Hooder Quality"] = value; break;
                case "gBoardQuality": _objectsService.MfgProcesscheck.dr["Board Quality"] = value; break;
            }
            _objectsService.MfgProcesscheck.UpdateDataSet();
            return new JsonResult(new { message = Name, value });
        }


        public IActionResult OnPostGMisc_LF(string Name, string value)
        {
           // Control ctl = sender as Control;
           // if (ctl == null) return null;
            //           DateTime _objectsService.MfgProcesscheck.dt;

            switch (Name)
            {
                case "gTestDate":
              _objectsService.MfgProcesscheck.dr["Sample Date Time"] = value;
                 break;
                case "gComment": _objectsService.MfgProcesscheck.dr["Comment"] = value ; break;
                case "gExclude": _objectsService.MfgProcesscheck.dr["bExclude"] = value; break;
                case "gExposedFoam": _objectsService.MfgProcesscheck.dr["Exposed Foam"]=value; break;

            }
            _objectsService.MfgProcesscheck.UpdateDataSet();
            return new JsonResult(new { message = Name, value });

        }
        public IActionResult OnPostGTestTime_LF(string Name, DateTime value)
        {

                if (_objectsService.MfgProcesscheck.dr["Sample Date Time"] == DBNull.Value) _objectsService.MfgProcesscheck.dr["Sample Date Time"] = ((DateTime)value);
                else _objectsService.MfgProcesscheck.dr["Sample Date Time"] = ((DateTime)_objectsService.MfgProcesscheck.dr["Sample Date Time"]).Date + ((DateTime)value).TimeOfDay;
            _objectsService.MfgProcesscheck.UpdateDataSet();
            return new JsonResult(new { message = Name, value });
        }
        public IActionResult OnPostGCopyDataSet()
        {
            string sdt1, sdt2; string sql;
            int itmp; DataTable dtCopy = null; SqlDataAdapter daCopy = null;

            if (gCopyData < 0) {
              //  MessageBox.Show("Choose an apprpriate time window", _objectsService.MfgProcesscheck.cbfile.sAppName); 
                return null; }
            DateTime dt1 = DateTime.Now;

            int isql = gCopyData;
            //            sql = "SELECT  * FROM [dbo].[Process Check] where IDLocation = " + _objectsService.MfgProcesscheck.CDefault.IDLocation.ToString();

            sql = "SELECT RN.ID, RN.[Product Code], RN.[Sample Date Time], R1.Employees, R2.sName as 'Check Type', R3.sName as 'Top Board Print', R4.sName as 'Bottom Board Print', R5.sName as 'Perferation', R6.sName as 'Flipper Operating', R7.sName as '[Facer Adhesion]', R8.sName as 'Edge Cut', R9.sName as 'Hooder Quality', R10.sName as 'Board Quality', R11.sName as 'Process Check Type', RN.Comment, RN.[Exposed Foam], RN.[Bundle Quantity 1] as 'Bundle 1 - Board Quantity', RN.[Bundle Wi_objectsService.MfgProcesscheck.dth 1] as 'Bundle 1 - Wi_objectsService.MfgProcesscheck.dth', RN.[Top Board Length 1] as 'Bundle 1 - Top Board Length', RN.[Middle Board Length 1] as 'Bundle 1 - Middle Board Length', RN.[Bottom Board Length 1] as 'Bundle 1 - Bottom Board Length', RN.[Diagonal_1 1] as 'Bundle 1 - Diagonal 1',RN.[Diagonal_2 1] as 'Bundle 1 - Diagonal 2', RN.[Length Average 1] as 'Bundle 1 - Average Length', RN.[Wi_objectsService.MfgProcesscheck.dth Average 1] as 'Bundle 1 - Average Wi_objectsService.MfgProcesscheck.dth', RN.[Squareness 1] as 'Bundle 1 - Squareness', RN.[Bundle Quantity 2] as 'Bundle 2 - Board Quantity', RN.[Bundle Wi_objectsService.MfgProcesscheck.dth 2] 'Bundle 2 - Wi_objectsService.MfgProcesscheck.dth', RN.[Top Board Length 2] as 'Bundle 2 - Top Board Length', RN.[Middle Board Length 2] as 'Bundle 2 - Middle Board Length', RN.[Bottom Board Length 2] as 'Bundle 2 - Bottom Board Length', RN.[Diagonal_1 2] as 'Bundle 2 - Diagonal 1', RN.[Diagonal_2 2] as 'Bundle 2 - Diagonal 2', RN.[Length Average 2] as 'Bundle 2 - Average Length', RN.[Wi_objectsService.MfgProcesscheck.dth Average 2] as 'Bundle 2 - Average Wi_objectsService.MfgProcesscheck.dth', RN.[Squareness 2] as 'Bundle 2 - Squareness', RN.ThicknessLoc1 as 'Board Thickness Location 1', RN.ThicknessLoc2 as 'Board Thickness Location 2', RN.ThicknessLoc3 as 'Board Thickness Location 3', RN.[Thickness Average] as 'Board Thickness Average', RN.Taper as 'Board Taper', R12.sLocation as 'Location', case when RN.bExclude = 1 then 'true' else 'false' end as 'Excluded from Anlysis if 1' " +
                "FROM[dbo].[Process Check] as RN Left Join[Roster] as R1 on RN.Operator = R1.ID Left Join tblLists as R2 on RN.[Check Type] = R2.ID Left Join tblLists as R3 on RN.[Top Board Print] = R3.ID Left Join tblLists as R4 on RN.[Bottom Board Print] = R4.ID Left Join tblLists as R5 on RN.Perferation = R5.ID Left Join tblLists as R6 on RN.[Flipper Operating] = R6.ID Left Join tblLists as R7 on RN.[Facer Adhesion] = R7.ID Left Join tblLists as R8 on RN.[Edge Cut] = R8.ID Left Join tblLists as R9 on RN.[Hooder Quality] = R9.ID Left Join tblLists as R10 on RN.[Board Quality] = R10.ID Left Join tblLists as R11 on RN.[Process Check Type] = R11.ID Left Join tblLocations as R12 on RN.IDLocation = R12.ID ";


            switch (isql)
            {
                case 0:dt1 =   dt1.Date; sql += " And [Sample Date Time] >= '" + dt1.ToString() + "' And [Sample Date Time] < '" + dt1.AddDays(1).ToString() + "'"; break;
                case 1: sql += " And [Sample Date Time] <= '" + dt1.ToString() + "' And [Sample Date Time] > '" + dt1.AddDays(-1).ToString() + "'"; break;
                case 2: sql += " And [Sample Date Time] <= '" + dt1.ToString() + "' And [Sample Date Time] > '" + dt1.AddDays(-7).ToString() + "'"; break;
                case 3: sql += " And [Sample Date Time] <= '" + dt1.ToString() + "' And [Sample Date Time] > '" + dt1.AddMonths(-1).ToString() + "'"; break;
                case 4: sql += " And [Sample Date Time] <= '" + dt1.ToString() + "' And [Sample Date Time] > '" + dt1.AddMonths(-6).ToString() + "'"; break;
                case 5: sql += " And [Sample Date Time] <= '" + dt1.ToString() + "' And [Sample Date Time] > '" + dt1.AddYears(-1).ToString() + "'"; break;
                case 6: break;
                default: return null;
            }
            sql += " and rn.idlocation = " + _objectsService.MfgProcesscheck.CDefault.IDLocation.ToString() + " order by [Sample Date Time] Desc  ";
           // Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                _objectsService.MfgProcesscheck.cbfile.conAZ.Open();
                daCopy = new SqlDataAdapter(sql, _objectsService.MfgProcesscheck.cbfile.conAZ);
                if (dtCopy == null) dtCopy = new DataTable(); else dtCopy.Reset();
                itmp = _objectsService.MfgProcesscheck.da.Fill(dtCopy);
                if (itmp < 1)
                {
                   // sMsgData = "There is no Process Check Data for you " + _objectsService.MfgProcesscheck.CDefault.sLocation + "during the selected time frame";
                    return null;

                }
            }
            catch (Exception ex)
            {
               // sMsgData = "Error in contacting Process Check Data for copying to clipboard ";
              //  MessageBox.Show(sMsgData, _objectsService.MfgProcesscheck.cbfile.sAppName);
               // System.Diagnostics.Trace.TraceError(sMsgData);
              //  CTelClient.TelException(ex, sMsgData);  //Azue Insight Trace Message
                return null;
            }
            finally
            {
                _objectsService.MfgProcesscheck.cbfile.conAZ.Close();
            }

            var sData = new StringBuilder();

            sData.Append(dtCopy.Columns[0].ColumnName.ToString());
            for (int icol = 1; icol < dtCopy.Columns.Count; icol++) sData.Append("\t" +dtCopy.Columns[icol].ColumnName.ToString());
            for (int irow = 0; irow <dtCopy.Rows.Count; irow++)
            {
                sData.Append("\n" +dtCopy.Rows[irow][0].ToString());
                for (int icol = 1; icol < dtCopy.Columns.Count; icol++) sData.Append("\t" + (dtCopy.Rows[irow][icol]).ToString());
            }
          //  Clipboard.SetText(sData.ToString());
            // Mouse.OverrideCursor = null;
            // CStatusBar.SetText("Search results copied to clipboard at " + DateTime.Now.ToString("hh:mm:ss:tt"));
            return new JsonResult(new { data = sData.ToString() });
        }
        public IActionResult OnPostDataSetNavi_Click(string direction)
        {
            string sFld = String.Empty;

            switch (direction)
            {
                case "prev": _objectsService.MfgProcesscheck.drIndex += 1; break;
                case "next": _objectsService.MfgProcesscheck.drIndex -= 1; break;
            }

            if (_objectsService.MfgProcesscheck.drIndex < 0) 
                _objectsService.MfgProcesscheck.drIndex = 0;
            if (_objectsService.MfgProcesscheck.drIndex > _objectsService.MfgProcesscheck.dt.Rows.Count - 1)
                _objectsService.MfgProcesscheck.drIndex = _objectsService.MfgProcesscheck.dt.Rows.Count - 1;

            _objectsService.MfgProcesscheck.dr = _objectsService.MfgProcesscheck.dt.Rows[_objectsService.MfgProcesscheck.drIndex];
            
            _objectsService.MfgProcesscheck.UpdateDataSet();
            SetView();
            return new JsonResult(new { message = direction,gID = _objectsService.MfgProcesscheck.dr["ID"] });
        }
    }
}
