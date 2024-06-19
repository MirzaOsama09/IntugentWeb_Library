using IntugentClassLbrary.Classes;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Runtime.Serialization;
using System;
using IntugentWebApp.Utilities;
using IntugentClassLibrary.Pages.Mfg;
using System.Xml.Linq;
using System.Globalization;
using static System.Net.Mime.MediaTypeNames;

namespace IntugentWebApp.Pages.Mfg_Group
{
    [BindProperties]
    public class FinishedBoardPropertiesModel : PageModel
    {
        public string gID { get; set; }
        public string gIPLength { get; set; }
        public string gIPWidth { get; set; }
        public string gTestDate { get; set; }
        public string gProdCode { get; set; }

        public string gBundleHeight { get; set; }
        public bool gTestingPassed { get; set; }
        public bool gFinsihedGoodsDone { get; set; }
        public string gTimePourTableQC { get; set; }
        public string gIPBoardTimeStamp { get; set; }
        public bool gIPTimeNotLegible { get; set; }
        public string gCoreDensity { get; set; }
        public string gCompStrFG_Avg6 { get; set; }
        public string gCompStrFG_Avg5 { get; set; }
        public string gThickness { get; set; }
        public string gFlatness { get; set; }
        public string gRValue { get; set; }
        public string gFacerPeelAvg_QC { get; set; }
        public string gWarehouseTemp { get; set; }
        public string gWarehouseHumidity { get; set; }
        public string gBundleHeightIP { get; set; }
        public string gThicknessIP { get; set; }
        public string gCoreDensityIP { get; set; }
        public string gCompressiveIP5 { get; set; }
        public string gCompressiveIP { get; set; }
        public string gThicknessFG_1 { get; set; }
        public string gThicknessFG_2 { get; set; }
        public string gThicknessFG_3 { get; set; }
        public string gThicknessFG_4 { get; set; }
        public string gThicknessFG_5 { get; set; }
        public string gThicknessFG_6 { get; set; }
        public string gThicknessFG_7 { get; set; }
        public string gThicknessFG_8 { get; set; }
        public string gThicknessFG_9 { get; set; }
        public string gThicknessFG_10 { get; set; }
        public string gThicknessFG_11 { get; set; }
        public string gThicknessFG_12 { get; set; }
        public string gThicknessFG_13 { get; set; }
        public string gThicknessFG_14 { get; set; }
        public string gThicknessFG_15 { get; set; }
        public string gThicknessFG_16 { get; set; }
        public string gThicknessFG_17 { get; set; }
        public string gLoc1F { get; set; }
        public string gLoc3F { get; set; }
        public string gCompStrFG_1 { get; set; }
        public string gCompStrFG_2 { get; set; }
        public string gCompStrFG_3 { get; set; }
        public string gCompStrFG_4 { get; set; }
        public string gCompStrFG_5 { get; set; }
        public string gCompStrFG_6 { get; set; }
        public bool gCompStrFGKnit_1 { get; set; }
        public bool gCompStrFGKnit_2 { get; set; }
        public bool gCompStrFGKnit_3 { get; set; }
        public bool gCompStrFGKnit_4 { get; set; }
        public bool gCompStrFGKnit_5 { get; set; }
        public bool gCompStrFGKnit_6 { get; set; }
        public string gNotes { get; set; }
        public bool gRestestFromSameBundle { get; set; }
        public string gCompStrFGRetest_1 { get; set; }
        public string gCompStrFGRetest_2 { get; set; }
        public string gCompStrFGRetest_3 { get; set; }
        public string gCompStrFGRetest_4 { get; set; }
        public string gCompStrFGRetest_5 { get; set; }
        public string gCompStrFGRetest_6 { get; set; }
        public bool gCompStrFGKnitRetest_1 { get; set; }
        public bool gCompStrFGKnitRetest_2 { get; set; }
        public bool gCompStrFGKnitRetest_3 { get; set; }
        public bool gCompStrFGKnitRetest_4 { get; set; }
        public bool gCompStrFGKnitRetest_5 { get; set; }
        public bool gCompStrFGKnitRetest_6 { get; set; }
        public string gCompStrFGRetest_Avg5 { get; set; }
        public string gCompStrFGRetest_Avg6 { get; set; }
        public string gLoc3B { get; set; }
        public string gLoc2B { get; set; }
        public string gLoc3C { get; set; }
        public string gLoc2C { get; set; }
        public string gLoc1C { get; set; }
        public string gkFactor90_3 { get; set; }
        public string gkFactor90_2 { get; set; }
        public string gkFactor90_1 { get; set; }
        public string gkFactor90 { get; set; }
        public string gLoc1B { get; set; }
        public string gkFactor_3 { get; set; }
        public string gkFactor_2 { get; set; }
        public string gkFactor_1 { get; set; }
        public string gkFactor180_3 { get; set; }
        public string gkFactor180_2 { get; set; }
        public string gkFactor180_1 { get; set; }
        public string gkFactor180 { get; set; }
        public bool gAgedRValueDone { get; set; }
        public string gLoggerID { get; set; }
        public string gInitProbeTemp { get; set; }
        public string gMaxProbeTemp { get; set; }
        public string gFinalProbeTemp { get; set; }
        public string gLoc3A { get; set; }
        public string gLoc2A { get; set; }
        public string gLoc1A { get; set; }
        public string gMass3 { get; set; }
        public string gMass2 { get; set; }
        public string gMass1 { get; set; }
        public string gL1_3 { get; set; }
        public string gL1_2 { get; set; }
        public string gL1_1 { get; set; }
        public string gW1_3 { get; set; }
        public string gW1_2 { get; set; }
        public string gW1_1 { get; set; }
        public string gT1_3 { get; set; }
        public string gT1_2 { get; set; }
        public string gT1_1 { get; set; }
        public string gT2_3 { get; set; }
        public string gT2_2 { get; set; }
        public string gT2_1 { get; set; }
        public string gT3_3 { get; set; }
        public string gT3_2 { get; set; }
        public string gT3_1 { get; set; }
        public string gT4_3 { get; set; }
        public string gT4_2 { get; set; }
        public string gT4_1 { get; set; }
        public string gT5_3 { get; set; }
        public string gT5_2 { get; set; }
        public string gT5_1 { get; set; }
        public string gAvgR { get; set; }
        public string gkFactor_Avg { get; set; }
        public bool gRValueKnitPresent3 { get; set; }
        public bool gRValueKnitPresent2 { get; set; }
        public bool gRValueKnitPresent1 { get; set; }
        public string gFGLength { get; set; }
        public string gFGWidth { get; set; }
        public string gFGDiagoanl1 { get; set; }
        public string gFGDiagoanl2 { get; set; }
        public string gFGDiagonalDiff { get; set; }
        public bool gCoreDensKnitLine3 { get; set; }
        public bool gCoreDensKnitLine2 { get; set; }
        public bool gCoreDensKnitLine1 { get; set; }
        public string gCoreDens3 { get; set; }
        public string gCoreDens2 { get; set; }
        public string gCoreDens1 { get; set; }
        public string gLoc3E { get; set; }
        public string gLoc2E { get; set; }
        public string gLoc1E { get; set; }
        public string gFacerPeel3 { get; set; }
        public string gFacerPeel2 { get; set; }
        public string gFacerPeel1 { get; set; }
        public string gFacerPeelAvg { get; set; }
        public string gLoc3D { get; set; }
        public string gLoc2D { get; set; }
        public string gLoc1D { get; set; }
        public string gNailPull_3 { get; set; }
        public string gNailPull_2 { get; set; }
        public string gNailPull_1 { get; set; }
        public string gNailPull { get; set; }
        //done properties
        public bool gDataSetPrev { get; set; }
        public bool gDataSetNext { get; set; }
        public DateTime? gKFactor90Date_3 { get; set; }
        public DateTime? gKFactor90Date_2 { get; set; }
        public DateTime? gKFactor90Date_1 { get; set; }
        public DateTime? gKFactor180Date_3 { get; set; }
        public DateTime? gKFactor180Date_2 { get; set; }
        public DateTime? gKFactor180Date_1 { get; set; }
        public DateTime? gFBTimeStamp { get; set; }
        public DateTime? gQCTimesDateTime { get; set; }
        public DateTime? gkFactorTime1 { get; set; }
        public DateTime? gkFactorTime2 { get; set; }
        public DateTime? gkFactorTime3 { get; set; }
        public DateTime? gInitProbeTime { get; set; }
        public DateTime? gMaxTempTimeInit { get; set; }
        public DateTime? gMaxTempTimeFinal { get; set; }
        public DateTime? gFinalProbeTime { get; set; }
        public DateTime? gRetestQCTime { get; set; }
        //public DateTime?  {get;set; }

        private readonly ObjectsService _objectsService;
        public FinishedBoardPropertiesModel(ObjectsService objectsService)
        {
            _objectsService = objectsService;
        }
        public void OnGet()
        {
            _objectsService.MfgFinishedGoods.GetDataSet();
            SetData();
        }
        public void SetData()
        {
            string stmp;
            int itmp;
            if (_objectsService.Cbfile.iIDMfgIndex == 0) gDataSetNext = false; else gDataSetNext = true;
            if (_objectsService.Cbfile.iIDMfgIndex == _objectsService.MfgHome.dt.Rows.Count - 1) gDataSetPrev = false; else gDataSetPrev = true;

            //           GetDataSet();
            //           CPages.PageInProcess_1.GetDataSet();
            _objectsService.MfgFinishedGoods.drIP = _objectsService.MfgInProcess.dr;

            gLoc1A = _objectsService.CDefualts.sLocMfg1.ToUpper();
            gLoc2A = _objectsService.CDefualts.sLocMfg2.ToUpper();
            gLoc3A = _objectsService.CDefualts.sLocMfg3.ToUpper();

            gLoc1B = gLoc1C = gLoc1D = gLoc1E = gLoc1F = _objectsService.CDefualts.sLocMfg1;
            gLoc2B = gLoc2C = gLoc2D = gLoc2E = _objectsService.CDefualts.sLocMfg2;
            gLoc3B = gLoc3C = gLoc3D = gLoc3E = gLoc3F = _objectsService.CDefualts.sLocMfg3;

            #region time, datetime, and date values

            if (_objectsService.MfgFinishedGoods.dr["Finished Board Time Stamp FG"] == DBNull.Value) gFBTimeStamp = null;
            else { gFBTimeStamp = (DateTime)_objectsService.MfgFinishedGoods.dr["Finished Board Time Stamp FG"]; }

            if (_objectsService.MfgFinishedGoods.dr["Next Day QC Collection Time FG"] == DBNull.Value) gQCTimesDateTime = null;
            else { gQCTimesDateTime = (DateTime)_objectsService.MfgFinishedGoods.dr["Next Day QC Collection Time FG"]; }

            if (_objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 1"] == DBNull.Value) gkFactorTime1 = null;
            else { gkFactorTime1 = (DateTime)_objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 1"]; }

            if (_objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 2"] == DBNull.Value) gkFactorTime2 = null;
            else { gkFactorTime2 = (DateTime)_objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 2"]; }

            if (_objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 3"] == DBNull.Value) gkFactorTime3 = null;
            else { gkFactorTime3 = (DateTime)_objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 3"]; }

            if (_objectsService.MfgFinishedGoods.dr["Initial Probe Time FG"] == DBNull.Value) gInitProbeTime = null;
            else { gInitProbeTime = (DateTime)_objectsService.MfgFinishedGoods.dr["Initial Probe Time FG"]; }

            if (_objectsService.MfgFinishedGoods.dr["Max Probe Time - Initial FG"] == DBNull.Value) gMaxTempTimeInit = null;
            else { gMaxTempTimeInit = (DateTime)_objectsService.MfgFinishedGoods.dr["Max Probe Time - Initial FG"]; }

            if (_objectsService.MfgFinishedGoods.dr["Max Probe Time - Final FG"] == DBNull.Value) gMaxTempTimeFinal = null;
            else { gMaxTempTimeFinal = (DateTime)_objectsService.MfgFinishedGoods.dr["Max Probe Time - Final FG"]; }

            if (_objectsService.MfgFinishedGoods.dr["Final Probe Time FG"] == DBNull.Value) gFinalProbeTime = null;
            else { gFinalProbeTime = (DateTime)_objectsService.MfgFinishedGoods.dr["Final Probe Time FG"]; }

            if (_objectsService.MfgFinishedGoods.dr["Retest QC Collection Time FG"] == DBNull.Value) gRetestQCTime = null;
            else { gRetestQCTime = (DateTime)_objectsService.MfgFinishedGoods.dr["Retest QC Collection Time FG"]; }

            //if (_objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 1"] == DBNull.Value) gKFactor90Date_1 = null;
            // else { gKFactor90Date_1 = (DateTime)_objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 1"]; }

            //SetDateTimeField("k Factor 90 Date FG 1");
            //SetDateTimeField("k Factor 90 Date FG 2");
            //SetDateTimeField("k Factor 90 Date FG 3");

            //SetDateTimeField("k Factor 180 Date FG 1");
            //SetDateTimeField("k Factor 180 Date FG 2");
            //SetDateTimeField("k Factor 180 Date FG 3");
            SetDateTimeField();

            #endregion

            #region General Info, QC Times, QC Summary, IP QC Summary

            gID = _objectsService.MfgFinishedGoods.dr["ID4ALL FG"].ToString();
            if (_objectsService.MfgFinishedGoods.dr["FG Testing Complete"] == DBNull.Value) gFinsihedGoodsDone = false; else gFinsihedGoodsDone = (bool)_objectsService.MfgFinishedGoods.dr["FG Testing Complete"];
            if (_objectsService.MfgFinishedGoods.dr["QC Test Passed"] == DBNull.Value) gTestingPassed = false; else gTestingPassed = (bool)_objectsService.MfgFinishedGoods.dr["QC Test Passed"];
            if (_objectsService.MfgFinishedGoods.drIP["Product ID"] == DBNull.Value) gProdCode = String.Empty;
            else
            {
                stmp = _objectsService.MfgFinishedGoods.drIP["Product ID"].ToString();
                _objectsService.CLists.dvComProd.RowFilter = "[Product Code] = '" + stmp + "'";
                DataTable dtxxx = _objectsService.CLists.dvComProd.ToTable();
                if (dtxxx.Rows.Count > 0 && dtxxx.Rows[0]["Product"] != DBNull.Value) gProdCode = dtxxx.Rows[0]["Product"].ToString();
                else gProdCode = string.Empty;
                _objectsService.CLists.dvComProd.RowFilter = null;
            }
            if (_objectsService.MfgFinishedGoods.drIP["Test Date"] == DBNull.Value) gTestDate = String.Empty; else gTestDate = ((DateTime)_objectsService.MfgFinishedGoods.drIP["Test Date"]).ToString("yyyy-MM-ddTHH:mm:ss");
            //            gFBTimeStamp.Value = SetDateTimeField("Finished Board Time Stamp FG");

            gBundleHeight = _objectsService.MfgFinishedGoods.SetDoubleTextField("Bundle Height FG");
            if (_objectsService.MfgFinishedGoods.drIP["Length"] == DBNull.Value) gIPLength = String.Empty; else gIPLength = ((double)_objectsService.MfgFinishedGoods.drIP["Length"]).ToString(MfgFinishedGoods.sOr);
            if (_objectsService.MfgFinishedGoods.drIP["Width"] == DBNull.Value) gIPWidth = String.Empty; else gIPWidth = ((double)_objectsService.MfgFinishedGoods.drIP["Width"]).ToString(MfgFinishedGoods.sOr);
            //          if (_objectsService.MfgFinishedGoods.drIP["IP Diagonal 1"] == DBNull.Value) gDiagoanl1.Text = String.Empty; else gDiagoanl1.Text = ((double)_objectsService.MfgFinishedGoods.drIP["IP Diagonal 1"]).ToString(MfgFinishedGoods.sOr);
            //          if (_objectsService.MfgFinishedGoods.drIP["IP Diagonal 2"] == DBNull.Value) gDiagoanl2.Text = String.Empty; else gDiagoanl2.Text = ((double)_objectsService.MfgFinishedGoods.drIP["IP Diagonal 2"]).ToString(MfgFinishedGoods.sOr);

            if (_objectsService.MfgFinishedGoods.drIP["Time of Pour Table QC Check"] == DBNull.Value) gTimePourTableQC = String.Empty; else gTimePourTableQC = ((DateTime)_objectsService.MfgFinishedGoods.drIP["Time of Pour Table QC Check"]).ToString("HH:mm:ss");
            if (_objectsService.MfgFinishedGoods.drIP["Test Date"] == DBNull.Value) gIPBoardTimeStamp = String.Empty; else gIPBoardTimeStamp = ((DateTime)_objectsService.MfgFinishedGoods.drIP["Test Date"]).ToString("yyyy-MM-ddTHH:mm:ss");
            //           gQCTimesDateTime.Value = SetDateTimeField("Next Day QC Collection Time FG");
            if (_objectsService.MfgFinishedGoods.dr["IP Time Stamp Not Legible"] == DBNull.Value) gIPTimeNotLegible = false; else gIPTimeNotLegible = (bool)_objectsService.MfgFinishedGoods.dr["IP Time Stamp Not Legible"];

            gCoreDensity = _objectsService.MfgFinishedGoods.SetDoubleTextField("FG Core Density", MfgFinishedGoods.sOr);
            gCompStrFG_Avg6 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Compressive Strength (6) FG", MfgFinishedGoods.sOr);
            gCompStrFG_Avg5 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Compressive Strength (5) FG", MfgFinishedGoods.sOr);
            gThickness = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness Avg FG", MfgFinishedGoods.sOr);
            gFlatness = _objectsService.MfgFinishedGoods.SetDoubleTextField("Flatness FG", MfgFinishedGoods.sOr);
            gRValue = _objectsService.MfgFinishedGoods.SetDoubleTextField("R Value - AVG FG", MfgFinishedGoods.sOr);
            gFacerPeelAvg_QC = _objectsService.MfgFinishedGoods.SetDoubleTextField("Facer Peel FG", MfgFinishedGoods.sOr);

            if (_objectsService.MfgFinishedGoods.drIP["Warehouse Temp"] == DBNull.Value) gWarehouseTemp = String.Empty; else gWarehouseTemp = ((double)_objectsService.MfgFinishedGoods.drIP["Warehouse Temp"]).ToString(MfgFinishedGoods.sOr);
            if (_objectsService.MfgFinishedGoods.drIP["Warehouse Humidity"] == DBNull.Value) gWarehouseHumidity = String.Empty; else gWarehouseHumidity = ((double)_objectsService.MfgFinishedGoods.drIP["Warehouse Humidity"]).ToString(MfgFinishedGoods.sOr);
            if (_objectsService.MfgFinishedGoods.drIP["Bundle Height IP"] == DBNull.Value) gBundleHeightIP = String.Empty; else gBundleHeightIP = ((double)_objectsService.MfgFinishedGoods.drIP["Bundle Height IP"]).ToString(MfgFinishedGoods.sOr);
            if (_objectsService.MfgFinishedGoods.drIP["Thickness - IP"] == DBNull.Value) gThicknessIP = String.Empty; else gThicknessIP = ((double)_objectsService.MfgFinishedGoods.drIP["Thickness - IP"]).ToString(MfgFinishedGoods.sOr);
            if (_objectsService.MfgFinishedGoods.drIP["Core Density - IP"] == DBNull.Value) gCoreDensityIP = String.Empty; else gCoreDensityIP = ((double)_objectsService.MfgFinishedGoods.drIP["Core Density - IP"]).ToString(MfgFinishedGoods.sOr);
            if (_objectsService.MfgFinishedGoods.drIP["Compressive Strength - IP"] == DBNull.Value) gCompressiveIP = String.Empty; else gCompressiveIP = ((double)_objectsService.MfgFinishedGoods.drIP["Compressive Strength - IP"]).ToString(MfgFinishedGoods.sOr);
            //            if (_objectsService.MfgFinishedGoods.drIP["IP Diagonal Diff"] == DBNull.Value) gDiagoanlDiff.Text = String.Empty; else gDiagoanlDiff.Text = ((double)_objectsService.MfgFinishedGoods.drIP["IP Diagonal Diff"]).ToString(MfgFinishedGoods.sOr);
            if (_objectsService.MfgFinishedGoods.drIP["Compressive Strength 5 - IP"] == DBNull.Value) gCompressiveIP5 = String.Empty; else gCompressiveIP5 = ((double)_objectsService.MfgFinishedGoods.drIP["Compressive Strength 5 - IP"]).ToString(MfgFinishedGoods.sOr);

            gFGLength = _objectsService.MfgFinishedGoods.SetDoubleTextField("Length FG");
            gFGWidth = _objectsService.MfgFinishedGoods.SetDoubleTextField("Width FG");
            gFGDiagoanl1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Diagonal FG 1");
            gFGDiagoanl2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Diagonal FG 2");
            gFGDiagonalDiff = _objectsService.MfgFinishedGoods.SetDoubleTextField("Diagonal Diff FG", MfgFinishedGoods.sOr);



            #endregion

            #region Thickness

            gThicknessFG_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 1");
            gThicknessFG_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 2");
            gThicknessFG_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 3");
            gThicknessFG_4 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 4");
            gThicknessFG_5 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 5");
            gThicknessFG_6 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 6");
            gThicknessFG_7 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 7");
            gThicknessFG_8 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 8");
            gThicknessFG_9 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 9");
            gThicknessFG_10 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 10");
            gThicknessFG_11 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 11");
            gThicknessFG_12 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 12");
            gThicknessFG_13 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 13");
            gThicknessFG_14 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 14");
            gThicknessFG_15 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 15");
            gThicknessFG_16 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 16");
            gThicknessFG_17 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Thickness FG - 17");


            #endregion

            #region Comp Str

            gCompStrFG_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Compressive FG - 1");
            gCompStrFG_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Compressive FG - 2");
            gCompStrFG_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Compressive FG - 3");
            gCompStrFG_4 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Compressive FG - 4");
            gCompStrFG_5 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Compressive FG - 5");
            gCompStrFG_6 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Compressive FG - 6");

            if ((_objectsService.MfgFinishedGoods.dr["Comp 1 Knit Present FG"] == DBNull.Value)) gCompStrFGKnit_1 = false; else gCompStrFGKnit_1 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 1 Knit Present FG"];
            if ((_objectsService.MfgFinishedGoods.dr["Comp 2 Knit Present FG"] == DBNull.Value)) gCompStrFGKnit_2 = false; else gCompStrFGKnit_2 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 2 Knit Present FG"];
            if ((_objectsService.MfgFinishedGoods.dr["Comp 3 Knit Present FG"] == DBNull.Value)) gCompStrFGKnit_3 = false; else gCompStrFGKnit_3 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 3 Knit Present FG"];
            if ((_objectsService.MfgFinishedGoods.dr["Comp 4 Knit Present FG"] == DBNull.Value)) gCompStrFGKnit_4 = false; else gCompStrFGKnit_4 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 4 Knit Present FG"];
            if ((_objectsService.MfgFinishedGoods.dr["Comp 5 Knit Present FG"] == DBNull.Value)) gCompStrFGKnit_5 = false; else gCompStrFGKnit_5 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 5 Knit Present FG"];
            if ((_objectsService.MfgFinishedGoods.dr["Comp 6 Knit Present FG"] == DBNull.Value)) gCompStrFGKnit_6 = false; else gCompStrFGKnit_6 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 6 Knit Present FG"];

            if ((_objectsService.MfgFinishedGoods.dr["Comp 1 Retest Knit Present FG"] == DBNull.Value)) gCompStrFGKnitRetest_1 = false; else gCompStrFGKnitRetest_1 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 1 Retest Knit Present FG"];
            if ((_objectsService.MfgFinishedGoods.dr["Comp 2 Retest Knit Present FG"] == DBNull.Value)) gCompStrFGKnitRetest_2 = false; else gCompStrFGKnitRetest_2 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 2 Retest Knit Present FG"];
            if ((_objectsService.MfgFinishedGoods.dr["Comp 3 Retest Knit Present FG"] == DBNull.Value)) gCompStrFGKnitRetest_3 = false; else gCompStrFGKnitRetest_3 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 3 Retest Knit Present FG"];
            if ((_objectsService.MfgFinishedGoods.dr["Comp 4 Retest Knit Present FG"] == DBNull.Value)) gCompStrFGKnitRetest_4 = false; else gCompStrFGKnitRetest_4 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 4 Retest Knit Present FG"];
            if ((_objectsService.MfgFinishedGoods.dr["Comp 5 Retest Knit Present FG"] == DBNull.Value)) gCompStrFGKnitRetest_5 = false; else gCompStrFGKnitRetest_5 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 5 Retest Knit Present FG"];
            if ((_objectsService.MfgFinishedGoods.dr["Comp 6 Retest Knit Present FG"] == DBNull.Value)) gCompStrFGKnitRetest_6 = false; else gCompStrFGKnitRetest_6 = (bool)_objectsService.MfgFinishedGoods.dr["Comp 6 Retest Knit Present FG"];


            if ((_objectsService.MfgFinishedGoods.dr["Notes FG"] == DBNull.Value)) gNotes = String.Empty; else gNotes = ((string)_objectsService.MfgFinishedGoods.dr["Notes FG"]);

            gCompStrFGRetest_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Retest - Comp 1 FG");
            gCompStrFGRetest_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Retest - Comp 2 FG");
            gCompStrFGRetest_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Retest - Comp 3 FG");
            gCompStrFGRetest_4 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Retest - Comp 4 FG");
            gCompStrFGRetest_5 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Retest - Comp 5 FG");
            gCompStrFGRetest_6 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Retest - Comp 6 FG");
            gCompStrFGRetest_Avg5 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Retest - AVG Comp Strength (5) FG", MfgFinishedGoods.sOr);
            gCompStrFGRetest_Avg6 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Retest - AVG Comp Strength (6) FG", MfgFinishedGoods.sOr);


            if (_objectsService.MfgFinishedGoods.dr["Is Retest From Same Bundle? FG"] == DBNull.Value) gRestestFromSameBundle = false; else gRestestFromSameBundle = (bool)_objectsService.MfgFinishedGoods.dr["Is Retest From Same Bundle? FG"];


            #endregion

            #region k Factor, Aged R Value

            gkFactor_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("k Factor 1 FG");
            gkFactor_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("k Factor 2 FG");
            gkFactor_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("k Factor 3 FG");
            gkFactor_Avg = _objectsService.MfgFinishedGoods.SetDoubleTextField("R Value - AVG FG", MfgFinishedGoods.sOr);

            if (_objectsService.MfgFinishedGoods.dr["R Value - Knit Present FG 1"] == DBNull.Value) gRValueKnitPresent1 = false; else gRValueKnitPresent1 = (bool)_objectsService.MfgFinishedGoods.dr["R Value - Knit Present FG 1"];
            if (_objectsService.MfgFinishedGoods.dr["R Value - Knit Present FG 2"] == DBNull.Value) gRValueKnitPresent2 = false; else gRValueKnitPresent2 = (bool)_objectsService.MfgFinishedGoods.dr["R Value - Knit Present FG 2"];
            if (_objectsService.MfgFinishedGoods.dr["R Value - Knit Present FG 3"] == DBNull.Value) gRValueKnitPresent3 = false; else gRValueKnitPresent3 = (bool)_objectsService.MfgFinishedGoods.dr["R Value - Knit Present FG 3"];

            //            gkFactorTime1.Value = SetDateTimeField("k Factor DateTime FG 1");
            //            gkFactorTime2.Value = SetDateTimeField("k Factor DateTime FG 2");
            //            gkFactorTime3.Value = SetDateTimeField("k Factor DateTime FG 3");

            gkFactor90_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("k Factor 90 FG 1");
            gkFactor90_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("k Factor 90 FG 2");
            gkFactor90_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("k Factor 90 FG 3");
            gkFactor90 = _objectsService.MfgFinishedGoods.SetDoubleTextField("R Value 90 - AVG FG", MfgFinishedGoods.sOr);

            gkFactor180_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("k Factor 180 FG 1");
            gkFactor180_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("k Factor 180 FG 2");
            gkFactor180_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("k Factor 180 FG 3");
            gkFactor180 = _objectsService.MfgFinishedGoods.SetDoubleTextField("R Value 180 - AVG FG", MfgFinishedGoods.sOr);

            if (_objectsService.MfgFinishedGoods.dr["FG Aged R Value Complete"] == DBNull.Value) gAgedRValueDone = false; else gAgedRValueDone = (bool)_objectsService.MfgFinishedGoods.dr["FG Aged R Value Complete"];

            #endregion

            #region Facel Peel, Nail Pull, Bundel Temps

            gFacerPeel1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Facer Peel 1 FG");
            gFacerPeel2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Facer Peel 2 FG");
            gFacerPeel3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Facer Peel 3 FG");
            gFacerPeelAvg = _objectsService.MfgFinishedGoods.SetDoubleTextField("Facer Peel FG", MfgFinishedGoods.sOr);

            gNailPull_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Nail Pull FG 1");
            gNailPull_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Nail Pull FG 2");
            gNailPull_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Nail Pull FG 3");
            gNailPull = _objectsService.MfgFinishedGoods.SetDoubleTextField("Nail Pull FG", MfgFinishedGoods.sOr);

            gLoggerID = _objectsService.MfgFinishedGoods.SetIntTextField("Logger ID # FG");
            gInitProbeTemp = _objectsService.MfgFinishedGoods.SetDoubleTextField("Initial Probe Temp FG");
            //gInitProbeTime = SetDateTimeField("Initial Probe Time FG");
            gMaxProbeTemp = _objectsService.MfgFinishedGoods.SetDoubleTextField("Max Probe Temp FG");
            //gMaxTempTimeInit = SetDateTimeField("Max Probe Time - Initial FG");
            //gMaxTempTimeFinal = SetDateTimeField("Max Probe Time - Final FG");
            gFinalProbeTemp = _objectsService.MfgFinishedGoods.SetDoubleTextField("Final Probe Temp FG");
            //           gFinalProbeTime.Value = SetDateTimeField("Final Probe Time FG");

            #endregion


            #region core density

            gMass1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Mass 1 FG");
            gL1_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("L1 1 FG");
            //           gL2_1.Text =  _objectsService.MfgFinishedGoods.SetDoubleTextField("L2 1 FG");
            gW1_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("W1 1 FG");
            //            gW2_1.Text =  _objectsService.MfgFinishedGoods.SetDoubleTextField("W2 1 FG");
            gT1_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T1 1 FG");
            gT2_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T2 1 FG");
            gT3_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T3 1 FG");
            gT4_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T4 1 FG");
            gT5_1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T5 1 FG");
            gCoreDens1 = _objectsService.MfgFinishedGoods.SetDoubleTextField("FG Core Density 1", MfgFinishedGoods.sOr);
            if ((_objectsService.MfgFinishedGoods.dr["Core Knit Present FG 1"] == DBNull.Value)) gCoreDensKnitLine1 = false; else gCoreDensKnitLine1 = (bool)_objectsService.MfgFinishedGoods.dr["Core Knit Present FG 1"];



            gMass2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Mass 2 FG");
            gL1_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("L1 2 FG");
            //            gL2_2.Text =  _objectsService.MfgFinishedGoods.SetDoubleTextField("L2 2 FG");
            gW1_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("W1 2 FG");
            //            gW2_2.Text =  _objectsService.MfgFinishedGoods.SetDoubleTextField("W2 2 FG");
            gT1_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T1 2 FG");
            gT2_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T2 2 FG");
            gT3_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T3 2 FG");
            gT4_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T4 2 FG");
            gT5_2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T5 2 FG");
            gCoreDens2 = _objectsService.MfgFinishedGoods.SetDoubleTextField("FG Core Density 2", MfgFinishedGoods.sOr);
            if ((_objectsService.MfgFinishedGoods.dr["Core Knit Present FG 2"] == DBNull.Value)) gCoreDensKnitLine2 = false; else gCoreDensKnitLine2 = (bool)_objectsService.MfgFinishedGoods.dr["Core Knit Present FG 2"];

            gMass3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("Mass 3 FG");
            gL1_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("L1 3 FG");
            //            gL2_3.Text =  _objectsService.MfgFinishedGoods.SetDoubleTextField("L2 3 FG");
            gW1_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("W1 3 FG");
            //            gW2_3.Text =  _objectsService.MfgFinishedGoods.SetDoubleTextField("W2 3 FG");
            gT1_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T1 3 FG");
            gT2_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T2 3 FG");
            gT3_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T3 3 FG");
            gT4_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T4 3 FG");
            gT5_3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("T5 3 FG");
            gCoreDens3 = _objectsService.MfgFinishedGoods.SetDoubleTextField("FG Core Density 3", MfgFinishedGoods.sOr);
            if ((_objectsService.MfgFinishedGoods.dr["Core Knit Present FG 3"] == DBNull.Value)) gCoreDensKnitLine3 = false; else gCoreDensKnitLine3 = (bool)_objectsService.MfgFinishedGoods.dr["Core Knit Present FG 3"];

            #endregion



            CheckLimits("All");


        }
        public void SetDateTimeField()
        {
            gKFactor90Date_1 = _objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 1"] == DBNull.Value ? (DateTime?)null : (DateTime)_objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 1"];
            gKFactor90Date_2 = _objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 2"] == DBNull.Value ? (DateTime?)null : (DateTime)_objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 2"];
            gKFactor90Date_3 = _objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 3"] == DBNull.Value ? (DateTime?)null : (DateTime)_objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 3"];
            gKFactor180Date_1 = _objectsService.MfgFinishedGoods.dr["k Factor 180 Date FG 1"] == DBNull.Value ? (DateTime?)null : (DateTime)_objectsService.MfgFinishedGoods.dr["k Factor 180 Date FG 1"];
            gKFactor180Date_2 = _objectsService.MfgFinishedGoods.dr["k Factor 180 Date FG 2"] == DBNull.Value ? (DateTime?)null : (DateTime)_objectsService.MfgFinishedGoods.dr["k Factor 180 Date FG 2"];
            gKFactor180Date_3 = _objectsService.MfgFinishedGoods.dr["k Factor 180 Date FG 3"] == DBNull.Value ? (DateTime?)null : (DateTime)_objectsService.MfgFinishedGoods.dr["k Factor 180 Date FG 3"]; 
        }
        private void CheckLimits(string sF)
        {
            //Must be included in setview and   change products
            string sRet = string.Empty;


            //if (sF == "All") CProdTargets.GetProductTargets();


            //if (sF == "All")
            //{

            //    if (_objectsService.MfgFinishedGoods.drIP["Length"] == DBNull.Value) gIPLength.Background = backColorCal;
            //    else
            //    {
            //        sRet = CIPProdTargets.BoardLengthWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Length"]);
            //        //if (sRet == "Red") gIPLength.Background = backColorRed; else if (sRet == "Esc") gIPLength.Background = backColorEsc; else gIPLength.Background = backColorCal;
            //    }

            //    if (_objectsService.MfgFinishedGoods.drIP["Width"] == DBNull.Value) gIPWidth.Background = backColorCal;
            //    else
            //    {
            //        sRet = CIPProdTargets.BoardWidthWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Width"]);
            //       // if (sRet == "Red") gIPWidth.Background = backColorRed; else if (sRet == "Esc") gIPWidth.Background = backColorEsc; else gIPWidth.Background = backColorCal;
            //    }

            //    /*              if (_objectsService.MfgFinishedGoods.drIP["IP Diagonal Diff"] == DBNull.Value) gDiagoanlDiff.Background = backColorCal;
            //                  else
            //                  {
            //                      sRet = CIPProdTargets.BoardSquarenessWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["IP Diagonal Diff"]);
            //                      if (sRet == "Red") gDiagoanlDiff.Background = backColorRed; else if (sRet == "Esc") gDiagoanlDiff.Background = backColorEsc;  else gDiagoanlDiff.Background = backColor;
            //                  }
            //    */
            //    if (_objectsService.MfgFinishedGoods.drIP["Thickness - IP"] == DBNull.Value) gThicknessIP.Background = backColorCal;
            //    else
            //    {
            //        sRet = CIPProdTargets.ThicknessWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Thickness - IP"]);
            //       // if (sRet == "Red") gThicknessIP.Background = backColorRed; else if (sRet == "Esc") gThicknessIP.Background = backColorEsc; else gThicknessIP.Background = backColorCal;
            //    }

            //    if (_objectsService.MfgFinishedGoods.drIP["Core Density - IP"] == DBNull.Value) gCoreDensityIP.Background = backColorCal;
            //    else
            //    {
            //        sRet = CIPProdTargets.CoreDensityWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Core Density - IP"]);
            //       // if (sRet == "Red") gCoreDensityIP.Background = backColorRed; else if (sRet == "Esc") gCoreDensityIP.Background = backColorEsc; else gCoreDensityIP.Background = backColorCal;
            //    }

            //    if (_objectsService.MfgFinishedGoods.drIP["Compressive Strength - IP"] == DBNull.Value) gCompressiveIP.Background = backColorCal;
            //    else
            //    {
            //        sRet = CIPProdTargets.CompressionStrWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Compressive Strength - IP"]);
            //        //if (sRet == "Red") gCompressiveIP.Background = backColorRed; else if (sRet == "Esc") gCompressiveIP.Background = backColorEsc; else gCompressiveIP.Background = backColorCal;
            //    }

            //    if (_objectsService.MfgFinishedGoods.drIP["Compressive Strength 5 - IP"] == DBNull.Value) gCompressiveIP5.Background = backColorCal;
            //    else
            //    {
            //        sRet = CIPProdTargets.CompressionStrWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Compressive Strength 5 - IP"]);
            //       // if (sRet == "Red") gCompressiveIP5.Background = backColorRed; else if (sRet == "Esc") gCompressiveIP5.Background = backColorEsc; else gCompressiveIP5.Background = backColorCal;
            //    }


            /*
                            if (_objectsService.MfgFinishedGoods.drIP["Length"] == DBNull.Value) gIPLength.Background = backColorCal;
                            else if (CProdTargets.LengthWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Length"]) == "N") gIPLength.Background = backColorRed; else gIPLength.Background = backColorCal;

                            if (_objectsService.MfgFinishedGoods.drIP["Width"] == DBNull.Value) gIPWidth.Background = backColorCal;
                            else if (CProdTargets.WidthWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Width"]) == "N") gIPWidth.Background = backColorRed; else gIPWidth.Background = backColorCal;

                            if (_objectsService.MfgFinishedGoods.drIP["Thickness - IP"] == DBNull.Value) gThicknessIP.Background = backColorCal;
                            else if (CProdTargets.ThicknessWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Thickness - IP"]) == "N") gThicknessIP.Background = backColorRed; else gThicknessIP.Background = backColorCal;

                            if (_objectsService.MfgFinishedGoods.drIP["Core Density - IP"] == DBNull.Value) gCoreDensityIP.Background = backColorCal;
                            else if (CProdTargets.CoreDensWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Core Density - IP"]) == "N") gCoreDensityIP.Background = backColorRed; else gCoreDensityIP.Background = backColorCal;

                            if (_objectsService.MfgFinishedGoods.drIP["Compressive Strength - IP"] == DBNull.Value) gCompressiveIP.Background = backColorCal;
                            else if (CProdTargets.CompressionWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Compressive Strength - IP"]) == "N") gCompressiveIP.Background = backColorRed; else gCompressiveIP.Background = backColorCal;

                            if (_objectsService.MfgFinishedGoods.drIP["Compressive Strength 5 - IP"] == DBNull.Value) gCompressiveIP5.Background = backColorCal;
                            else if (CProdTargets.CompressionWithinLimits((double)_objectsService.MfgFinishedGoods.drIP["Compressive Strength 5 - IP"]) == "N") gCompressiveIP5.Background = backColorRed; else gCompressiveIP5.Background = backColorCal;
            */
            //}

            //if (sF == "gCoreDensity" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["FG Core Density"] == DBNull.Value) gCoreDensity.Background = backColorCal;
            //    else if (CProdTargets.CoreDensWithinLimits((double)_objectsService.MfgFinishedGoods.dr["FG Core Density"]) == "N") gCoreDensity.Background = backColorRed; else gCoreDensity.Background = backColorCal;
            //}

            //if (sF == "gCoreDens1" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["FG Core Density 1"] == DBNull.Value) gCoreDens1.Background = backColorCal;
            //    else if (CProdTargets.CoreDensWithinLimits((double)_objectsService.MfgFinishedGoods.dr["FG Core Density 1"]) == "N") gCoreDens1.Background = backColorRed; else gCoreDens1.Background = backColorCal;
            //}
            //if (sF == "gCoreDens2" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["FG Core Density 2"] == DBNull.Value) gCoreDens2.Background = backColorCal;
            //    else if (CProdTargets.CoreDensWithinLimits((double)_objectsService.MfgFinishedGoods.dr["FG Core Density 2"]) == "N") gCoreDens2.Background = backColorRed; else gCoreDens2.Background = backColorCal;
            //}
            //if (sF == "gCoreDens3" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["FG Core Density 3"] == DBNull.Value) gCoreDens3.Background = backColorCal;
            //    else if (CProdTargets.CoreDensWithinLimits((double)_objectsService.MfgFinishedGoods.dr["FG Core Density 3"]) == "N") gCoreDens3.Background = backColorRed; else gCoreDens3.Background = backColorCal;
            //}

            //if (sF == "gCompStrFG_Avg6" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["Compressive Strength (6) FG"] == DBNull.Value) gCompStrFG_Avg6.Background = backColorCal;
            //    else if (CProdTargets.CompressionWithinLimits((double)_objectsService.MfgFinishedGoods.dr["Compressive Strength (6) FG"]) == "N") gCompStrFG_Avg6.Background = backColorRed; else gCompStrFG_Avg6.Background = backColorCal;
            //}

            //if (sF == "gCompStrFG_Avg5" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["Compressive Strength (5) FG"] == DBNull.Value) gCompStrFG_Avg5.Background = backColorCal;
            //    else if (CProdTargets.CompressionWithinLimits((double)_objectsService.MfgFinishedGoods.dr["Compressive Strength (5) FG"]) == "N") gCompStrFG_Avg5.Background = backColorRed; else gCompStrFG_Avg5.Background = backColorCal;
            //}

            //if (sF == "gThickness" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["Thickness Avg FG"] == DBNull.Value) gThickness.Background = backColorCal;
            //    else if (CProdTargets.ThicknessWithinLimits((double)_objectsService.MfgFinishedGoods.dr["Thickness Avg FG"]) == "N") gThickness.Background = backColorRed; else gThickness.Background = backColorCal;
            //}

            //if (sF == "gRValue" || sF == "gkFactor_Avg" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["R Value - AVG FG"] == DBNull.Value) gRValue.Background = gkFactor_Avg.Background = backColorCal;
            //    else if (CProdTargets.RValueAged90DWithinLimits((double)_objectsService.MfgFinishedGoods.dr["R Value - AVG FG"]) == "N") gRValue.Background = gkFactor_Avg.Background = backColorRed; else gRValue.Background = gkFactor_Avg.Background = backColorCal;
            //}

            //if (sF == "gFacerPeelAvg" || sF == "gFacerPeelAvg_QC" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["Facer Peel FG"] == DBNull.Value) gFacerPeelAvg.Background = gFacerPeelAvg_QC.Background = backColorCal;
            //    else if (CProdTargets.FacerPeelWithinLimits((double)_objectsService.MfgFinishedGoods.dr["Facer Peel FG"]) == "N") gFacerPeelAvg.Background = gFacerPeelAvg_QC.Background = backColorRed; else gFacerPeelAvg.Background = gFacerPeelAvg_QC.Background = backColorCal;
            //}

            //if (sF == "gCompStrFGRetest_Avg6" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["Retest - AVG Comp Strength (6) FG"] == DBNull.Value) gCompStrFGRetest_Avg6.Background = backColorCal;
            //    else if (CProdTargets.CompressionWithinLimits((double)_objectsService.MfgFinishedGoods.dr["Retest - AVG Comp Strength (6) FG"]) == "N") gCompStrFGRetest_Avg6.Background = backColorRed; else gCompStrFGRetest_Avg6.Background = backColorCal;
            //}

            //if (sF == "gCompStrFGRetest_Avg5" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["Retest - AVG Comp Strength (5) FG"] == DBNull.Value) gCompStrFGRetest_Avg5.Background = backColorCal;
            //    else if (CProdTargets.CompressionWithinLimits((double)_objectsService.MfgFinishedGoods.dr["Retest - AVG Comp Strength (5) FG"]) == "N") gCompStrFGRetest_Avg5.Background = backColorRed; else gCompStrFGRetest_Avg5.Background = backColorCal;
            //}

            //if (sF == "gkFactor90" || sF == "gkFactor_Avg" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["R Value 90 - AVG FG"] == DBNull.Value) gkFactor90.Background = backColorCal;
            //    else if (CProdTargets.RValueAged90DWithinLimits((double)_objectsService.MfgFinishedGoods.dr["R Value 90 - AVG FG"]) == "N") gkFactor90.Background = backColorRed; else gkFactor90.Background = backColorCal;
            //}

            //if (sF == "gkFactor180" || sF == "gkFactor_Avg" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["R Value 180 - AVG FG"] == DBNull.Value) gkFactor180.Background = backColorCal;
            //    else if (CProdTargets.RValueAged90DWithinLimits((double)_objectsService.MfgFinishedGoods.dr["R Value 180 - AVG FG"]) == "N") gkFactor180.Background = backColorRed; else gkFactor180.Background = backColorCal;
            //}

            //if (sF == "gFGDiagonalDiff" || sF == "All")
            //{
            //    if (_objectsService.MfgFinishedGoods.dr["Diagonal Diff FG"] == DBNull.Value) gFGDiagonalDiff.Background = backColorCal;
            //    else if (CProdTargets.SquarenessWithinLimits((double)_objectsService.MfgFinishedGoods.dr["Diagonal Diff FG"]) == "N") gFGDiagonalDiff.Background = backColorRed; else gFGDiagonalDiff.Background = backColorCal;
            //}

            DateTime IPdate, FGdate;
            string sMsg = "Green Board and FG Board time stamps must be greater than 01-01-2000 and must be within " + _objectsService.CDefualts.dDelTimeButton.ToString() + " minutes of each other.";

            if (sF == "BoardTimeStamp" || sF == "All")
            {
                if (_objectsService.MfgFinishedGoods.dr["Finished Board Time Stamp FG"] == DBNull.Value || _objectsService.MfgFinishedGoods.drIP["Test Date"] == DBNull.Value)
                {
                    //gFBTimeHost.Background = backColorRed; if (sF == "BoardTimeStamp")
                    //   MessageBox.Show(sMsg, _objectsService.Cbfile.sAppName); 
                }
                else
                {
                    FGdate = (DateTime)_objectsService.MfgFinishedGoods.dr["Finished Board Time Stamp FG"];
                    IPdate = (DateTime)_objectsService.MfgFinishedGoods.drIP["Test Date"];
                    if (FGdate < new DateTime(2000, 01, 01) || Math.Abs((FGdate - IPdate).TotalMinutes) > _objectsService.CDefualts.dDelTimeButton)
                    {
                        // gFBTimeHost.Background = backColorRed; 
                    }
                    else
                    //gFBTimeHost.Background = backColor;

                    if (sF == "BoardTimeStamp")
                    {
                        //                       if (gFBTimeHost.Background != backColorRed) CPages.PagePlantData_1.GetPlantDataBackground(FGdate);
                        //                       else MessageBox.Show(sMsg, Cbfile.sAppName);

                        //if (gFBTimeHost.Background == backColor)
                        // {
                        //    CStatusBar.SetText("Pulling process data for dataset " + _objectsService.Cbfile.iIDMfg.ToString());
                        //    CPages.PagePlantData_1.GetPlantDataBackground(FGdate);
                        //                           CStatusBar.SetText("Finished pulling process data for dataset " + Cbfile.iIDMfg.ToString());
                        // }
                        // else
                        // {

                        //MessageBox.Show(sMsg, Cbfile.sAppName);
                        //  }
                    }
                }
            }
        }
        public IActionResult OnPostNavigateDataSet(string direction)
        {
            if (!_objectsService.Cbfile.bCanSwitchRecord)
            {
                //   MessageBox.Show(Cbfile.sNoRecSwitchMsg, Cbfile.sAppName);
                return new JsonResult(new { success = false, message = "Cannot switch record" });

            }

            switch (direction)
            {
                case "prev":
                    _objectsService.Cbfile.iIDMfgIndex += 1;
                    break;
                case "next":
                    _objectsService.Cbfile.iIDMfgIndex -= 1;
                    break;
                default:
                    return new JsonResult(new { success = false, message = "Invalid direction" });
            }

            if (_objectsService.Cbfile.iIDMfgIndex < 0)
            {
                _objectsService.Cbfile.iIDMfgIndex = 0;
            }

            if (_objectsService.Cbfile.iIDMfgIndex > _objectsService.MfgHome.dt.Rows.Count - 1)
            {
                _objectsService.Cbfile.iIDMfgIndex = _objectsService.MfgHome.dt.Rows.Count - 1;
            }

            UpdateDataSetView();
            return new JsonResult(new
            {
                success = true,
                newIndex = _objectsService.Cbfile.iIDMfgIndex
            });
        }
        private void UpdateDataSetView()
        {
            _objectsService.Cbfile.iIDMfg = (int)_objectsService.MfgHome.dt.Rows[_objectsService.Cbfile.iIDMfgIndex]["ID4All"];
            _objectsService.CLists.drEmployee["MfgIDSelected"] = _objectsService.Cbfile.iIDMfg;
            _objectsService.CLists.UpdateEmployee();

            (_objectsService.MfgInProcess, _objectsService.MfgFinishedGoods,_objectsService.MfgDimensionsStability, _objectsService.MfgPlantsData) = _objectsService.MfgHome.GetAllMfgData(_objectsService.MfgInProcess, _objectsService.MfgFinishedGoods, _objectsService.MfgDimensionsStability, _objectsService.MfgPlantsData);
            SetData();

            // Enable/disable buttons based on the index
            gDataSetNext = _objectsService.Cbfile.iIDMfgIndex < (_objectsService.MfgHome.dt.Rows.Count - 1);
            gDataSetPrev = _objectsService.Cbfile.iIDMfgIndex > 0;
        }
        public IActionResult OnPostCheckBoxes_LF(string Name, string value)
        {
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            string sName = Name;
            string Value = value;

            switch (sName)
            {
                case "gCoreDensKnitLine1":
                    _objectsService.MfgFinishedGoods.dr["Core Knit Present FG 1"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
                    break;
                case "gCoreDensKnitLine2":
                    _objectsService.MfgFinishedGoods.dr["Core Knit Present FG 2"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
                    break;
                case "gCoreDensKnitLine3":
                    _objectsService.MfgFinishedGoods.dr["Core Knit Present FG 3"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
                    break;

                case "gRValueKnitPresent1":
                    _objectsService.MfgFinishedGoods.dr["R Value - Knit Present FG 1"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
                    break;
                case "gRValueKnitPresent2":
                    _objectsService.MfgFinishedGoods.dr["R Value - Knit Present FG 2"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
                    break;
                case "gRValueKnitPresent3":
                    _objectsService.MfgFinishedGoods.dr["R Value - Knit Present FG 3"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
                    break;

                case "gAgedrValueDone":
                    _objectsService.MfgFinishedGoods.dr["FG Aged R Value Complete"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
                    break;

                case "gFinsihedGoodsDone":
                    if (string.IsNullOrEmpty(value) || !bool.Parse(value) || !_objectsService.gInProcessDoneIsChecked)
                    {
                        _objectsService.MfgFinishedGoods.dr["FG Testing Complete"] = false;
                        value = "false";
                    }
                    else
                    {
                        _objectsService.MfgFinishedGoods.dr["FG Testing Complete"] = true;
                        value = "true";
                    }
                    break;

                case "gTestingPassed":
                    _objectsService.MfgFinishedGoods.dr["QC Test Passed"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
                    break;
            }

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { Message = $"{sName}: {value}" });
        }



        public IActionResult OnPostCompStrFG_LF(string Name, string Value)
        {
            string sFld = String.Empty;
            //TextBox tbx = sender as TextBox;
            //string sText = tbx.Text;
            //string sName = tbx.Name;
            string sName = Name;
            string sText = Value;
            switch (sName)
            {
                case "gCompStrFG_1": _objectsService.MfgFinishedGoods.dr["Compressive FG - 1"] = Value; break;
                case "gCompStrFG_2": _objectsService.MfgFinishedGoods.dr["Compressive FG - 2"] = Value; break;
                case "gCompStrFG_3": _objectsService.MfgFinishedGoods.dr["Compressive FG - 3"] = Value; break;
                case "gCompStrFG_4": _objectsService.MfgFinishedGoods.dr["Compressive FG - 4"] = Value; break;
                case "gCompStrFG_5": _objectsService.MfgFinishedGoods.dr["Compressive FG - 5"] = Value; break;
                case "gCompStrFG_6": _objectsService.MfgFinishedGoods.dr["Compressive FG - 6"] = Value; break;

            }

            int nCount = 0;
            double dSum = 0.0, dMin = double.MaxValue, dtmp;

            if (!(_objectsService.MfgFinishedGoods.dr["Compressive FG - 1"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Compressive FG - 1"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }
            if (!(_objectsService.MfgFinishedGoods.dr["Compressive FG - 2"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Compressive FG - 2"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }
            if (!(_objectsService.MfgFinishedGoods.dr["Compressive FG - 3"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Compressive FG - 3"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }
            if (!(_objectsService.MfgFinishedGoods.dr["Compressive FG - 4"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Compressive FG - 4"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }
            if (!(_objectsService.MfgFinishedGoods.dr["Compressive FG - 5"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Compressive FG - 5"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }
            if (!(_objectsService.MfgFinishedGoods.dr["Compressive FG - 6"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Compressive FG - 6"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }

            if (nCount == 6)
            {
                _objectsService.MfgFinishedGoods.dr["Compressive Strength (6) FG"] = dSum / 6.0; gCompStrFG_Avg6 = (dSum / 6.0).ToString(MfgFinishedGoods.sOr);
                _objectsService.MfgFinishedGoods.dr["Compressive Strength (5) FG"] = (dSum - dMin) / 5.0; gCompStrFG_Avg5 = ((dSum - dMin) / 5.0).ToString(MfgFinishedGoods.sOr);
            }
            else if (nCount == 5)
            {
                _objectsService.MfgFinishedGoods.dr["Compressive Strength (6) FG"] = DBNull.Value; gCompStrFG_Avg6 = String.Empty;
                _objectsService.MfgFinishedGoods.dr["Compressive Strength (5) FG"] = dSum / 5.0; gCompStrFG_Avg5 = (dSum / 5.0).ToString(MfgFinishedGoods.sOr);
            }

            else
            {
                _objectsService.MfgFinishedGoods.dr["Compressive Strength (6) FG"] = DBNull.Value; gCompStrFG_Avg6 = String.Empty;
                _objectsService.MfgFinishedGoods.dr["Compressive Strength (5) FG"] = DBNull.Value; gCompStrFG_Avg5 = String.Empty;
            }
            _objectsService.MfgFinishedGoods.UpdateDataSet();
            //CheckLimits("gCompStrFG_Avg6"); CheckLimits("gCompStrFG_Avg5");
            return new JsonResult(new { Message = Name, Value });
        }

    public IActionResult OnPostGCompStrFG2_LF(string Name,string Value)
{
    string sName = Name;
    string value = Value;

    _objectsService.MfgFinishedGoods.bDataSetChanged = true;

    switch (sName)
    {
        case "gNotes":
            if (string.IsNullOrEmpty(value))
                _objectsService.MfgFinishedGoods.dr["Notes FG"] = DBNull.Value;
            else
                _objectsService.MfgFinishedGoods.dr["Notes FG"] = value;
            break;
        case "gCompStrFGKnit_1":
            _objectsService.MfgFinishedGoods.dr["Comp 1 Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnit_2":
            _objectsService.MfgFinishedGoods.dr["Comp 2 Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnit_3":
            _objectsService.MfgFinishedGoods.dr["Comp 3 Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnit_4":
            _objectsService.MfgFinishedGoods.dr["Comp 4 Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnit_5":
            _objectsService.MfgFinishedGoods.dr["Comp 5 Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnit_6":
            _objectsService.MfgFinishedGoods.dr["Comp 6 Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnitRetest_1":
            _objectsService.MfgFinishedGoods.dr["Comp 1 Retest Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnitRetest_2":
            _objectsService.MfgFinishedGoods.dr["Comp 2 Retest Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnitRetest_3":
            _objectsService.MfgFinishedGoods.dr["Comp 3 Retest Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnitRetest_4":
            _objectsService.MfgFinishedGoods.dr["Comp 4 Retest Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnitRetest_5":
            _objectsService.MfgFinishedGoods.dr["Comp 5 Retest Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
        case "gCompStrFGKnitRetest_6":
            _objectsService.MfgFinishedGoods.dr["Comp 6 Retest Knit Present FG"] = string.IsNullOrEmpty(value) ? (object)DBNull.Value : (object)bool.Parse(value);
            break;
    }

    _objectsService.MfgFinishedGoods.UpdateDataSet();
    return new JsonResult(new { message = $"{sName}: {value}" });
}

        public IActionResult OnPostCompStrFGRetest_LF(string Name, string Value)
        {
            string Text = Value;
            string sFld = String.Empty;
            //TextBox tbx = sender as TextBox;
            //string sText = tbx.Text;
            //string sName = tbx.Name;
            string sName = Name;
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            switch (sName)
            {
                case "gCompStrFGRetest_1": _objectsService.MfgFinishedGoods.dr["Retest - Comp 1 FG"] = Text; break;
                case "gCompStrFGRetest_2": _objectsService.MfgFinishedGoods.dr["Retest - Comp 2 FG"] = Text; break;
                case "gCompStrFGRetest_3": _objectsService.MfgFinishedGoods.dr["Retest - Comp 3 FG"] = Text; break;
                case "gCompStrFGRetest_4": _objectsService.MfgFinishedGoods.dr["Retest - Comp 4 FG"] = Text; break;
                case "gCompStrFGRetest_5": _objectsService.MfgFinishedGoods.dr["Retest - Comp 5 FG"] = Text; break;
                case "gCompStrFGRetest_6": _objectsService.MfgFinishedGoods.dr["Retest - Comp 6 FG"] = Text; break;
            }

            int nCount = 0;
            double dSum = 0.0, dMin = double.MaxValue, dtmp;

            if (!(_objectsService.MfgFinishedGoods.dr["Retest - Comp 1 FG"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Retest - Comp 1 FG"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }
            if (!(_objectsService.MfgFinishedGoods.dr["Retest - Comp 2 FG"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Retest - Comp 2 FG"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }
            if (!(_objectsService.MfgFinishedGoods.dr["Retest - Comp 3 FG"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Retest - Comp 3 FG"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }
            if (!(_objectsService.MfgFinishedGoods.dr["Retest - Comp 4 FG"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Retest - Comp 4 FG"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }
            if (!(_objectsService.MfgFinishedGoods.dr["Retest - Comp 5 FG"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Retest - Comp 5 FG"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }
            if (!(_objectsService.MfgFinishedGoods.dr["Retest - Comp 6 FG"] == DBNull.Value)) { nCount += 1; dtmp = (double)_objectsService.MfgFinishedGoods.dr["Retest - Comp 6 FG"]; dSum += dtmp; if (dMin > dtmp) dMin = dtmp; }

            if (nCount == 6)
            {
                _objectsService.MfgFinishedGoods.dr["Retest - AVG Comp Strength (6) FG"] = dSum / 6.0; gCompStrFGRetest_Avg6 = (dSum / 6.0).ToString(MfgFinishedGoods.sOr);
                _objectsService.MfgFinishedGoods.dr["Retest - AVG Comp Strength (5) FG"] = (dSum - dMin) / 5.0; gCompStrFGRetest_Avg5 = ((dSum - dMin) / 5.0).ToString(MfgFinishedGoods.sOr);
            }
            else if (nCount == 5)
            {
                _objectsService.MfgFinishedGoods.dr["Retest - AVG Comp Strength (6) FG"] = DBNull.Value; gCompStrFGRetest_Avg6 = String.Empty;
                _objectsService.MfgFinishedGoods.dr["Retest - AVG Comp Strength (5) FG"] = dSum / 5.0; gCompStrFGRetest_Avg5 = (dSum / 5.0).ToString(MfgFinishedGoods.sOr);
            }

            else
            {
                _objectsService.MfgFinishedGoods.dr["Retest - AVG Comp Strength (6) FG"] = DBNull.Value; gCompStrFGRetest_Avg6 = String.Empty;
                _objectsService.MfgFinishedGoods.dr["Retest - AVG Comp Strength (5) FG"] = DBNull.Value; gCompStrFGRetest_Avg5 = String.Empty;
            }
            _objectsService.MfgFinishedGoods.UpdateDataSet();

            //CheckLimits("gCompStrFGRetest_Avg5"); CheckLimits("gCompStrFGRetest_Avg6");

            return new JsonResult(new { message = "Dataset selected: " + Name + " -- " + Text });
        }
        public IActionResult OnPostGRestestFromSameBundle_LF(string Value)
        {
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            bool? gRestestFromSameBundle = string.IsNullOrEmpty(Value) ? (bool?)null : bool.Parse(Value);

            if (gRestestFromSameBundle == null)
                _objectsService.MfgFinishedGoods.dr["Is Retest From Same Bundle? FG"] = DBNull.Value;
            else
                _objectsService.MfgFinishedGoods.dr["Is Retest From Same Bundle? FG"] = gRestestFromSameBundle;

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = Value});
        }

        public IActionResult OnPostGkFactor_LF(string Name, double value)
        {
            string sFld = String.Empty;
            string sName = Name;
            double sText = value;

            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            switch (sName)
            {
                case "gkFactor_1": _objectsService.MfgFinishedGoods.dr["k Factor 1 FG"] = sText; break;
                case "gkFactor_2": _objectsService.MfgFinishedGoods.dr["k Factor 2 FG"] = sText; break;
                case "gkFactor_3": _objectsService.MfgFinishedGoods.dr["k Factor 3 FG"] = sText; break;
            }


            int nCount = 0;
            double dSum = 0, dtmp, dtmpR, dSumR = 0;

            if (!(_objectsService.MfgFinishedGoods.dr["k Factor 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["k Factor 1 FG"]; dSumR += 1.0 / (double)_objectsService.MfgFinishedGoods.dr["k Factor 1 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["k Factor 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["k Factor 2 FG"]; dSumR += 1.0 / (double)_objectsService.MfgFinishedGoods.dr["k Factor 2 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["k Factor 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["k Factor 3 FG"]; dSumR += 1.0 / (double)_objectsService.MfgFinishedGoods.dr["k Factor 3 FG"]; }

            if (nCount > 1)
            {
                dtmp = (dSum / (double)nCount); _objectsService.MfgFinishedGoods.dr["k Factor FG"] = dtmp;
                dtmpR = (dSumR / (double)nCount); gRValue = gkFactor_Avg = dtmpR.ToString(MfgFinishedGoods.sOr); _objectsService.MfgFinishedGoods.dr["R Value - AVG FG"] = dtmpR;
            }
            else { gkFactor_Avg = gRValue = String.Empty; _objectsService.MfgFinishedGoods.dr["k Factor FG"] = _objectsService.MfgFinishedGoods.dr["R Value - AVG FG"] = DBNull.Value; }

            CheckLimits("gkFactor_Avg"); // CheckLimits("gRValue"); //gRValue is taken care with gkFactor_Avg
            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = "Dataset selected: " + sName + " -- " + sText });


            // Return JsonResult with JSON data
        }
        public IActionResult OnPostCoreDenAv1_LF(string Name, string value)
        {
            string sFld = String.Empty;
            string sText = value;
            string sName = Name;

            switch (sName)
            {
                case "gMass1": _objectsService.MfgFinishedGoods.dr["Mass 1 FG"] = value; break;
                case "gL1_1": _objectsService.MfgFinishedGoods.dr["L1 1 FG"] = value; break;
                //            _objectsService.MfgFinishedGoods.dr[oubleField=value(gL2_1, "L2 1 FG"); break;
                case "gW1_1": _objectsService.MfgFinishedGoods.dr["W1 1 FG"] = value; break;
                //            _objectsService.MfgFinishedGoods.dr[ubleField(=valuesText, "W2 1 FG"); break;
                case "gT1_1": _objectsService.MfgFinishedGoods.dr["T1 1 FG"] = value; break;
                case "gT2_1": _objectsService.MfgFinishedGoods.dr["T2 1 FG"] = value; break;
                case "gT3_1": _objectsService.MfgFinishedGoods.dr["T3 1 FG"] = value; break;
                case "gT4_1": _objectsService.MfgFinishedGoods.dr["T4 1 FG"] = value; break;
                case "gT5_1": _objectsService.MfgFinishedGoods.dr["T5 1 FG"] = value; break;

            }

            int nCount = 0; double dSum = 0, dtmp;
            double dm = 0, dl, dw, dt; int nm, nl, nw, nt;

            nm = nl = nw = nt = 0;
            dm = dl = dw = dt = 0.0;

            if (!(_objectsService.MfgFinishedGoods.dr["Mass 1 FG"] == DBNull.Value)) { dm = (double)_objectsService.MfgFinishedGoods.dr["Mass 1 FG"]; nm = 1; }

            nCount = 0; dSum = 0;
            if (!(_objectsService.MfgFinishedGoods.dr["L1 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["L1 1 FG"]; }
            //            if (!(dr["L2 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)dr["L2 1 FG"]; }
            if (nCount > 0) { dl = dSum / (double)nCount; nl = 2; }

            nCount = 0; dSum = 0;
            if (!(_objectsService.MfgFinishedGoods.dr["W1 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["W1 1 FG"]; }
            //            if (!(dr["W2 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)dr["W2 1 FG"]; }
            if (nCount > 0) { dw = dSum / (double)nCount; nw = 2; }

            nCount = 0; dSum = 0;
            if (!(_objectsService.MfgFinishedGoods.dr["T1 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T1 1 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T2 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T2 1 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T3 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T3 1 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T4 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T4 1 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T5 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T5 1 FG"]; }

            if (nCount > 3) { dt = dSum / (double)nCount; nt = 5; }

            if (nm * nl * nw * nt > 0 && dl * dw * dt > 0.0) { dtmp = dm / (dl * dw * dt) * 3.809590998; gCoreDens1 = dtmp.ToString(MfgFinishedGoods.sOr); _objectsService.MfgFinishedGoods.dr["FG Core Density 1"] = dtmp; }
            else { gCoreDens1 = String.Empty; _objectsService.MfgFinishedGoods.dr["FG Core Density 1"] = DBNull.Value; }

            CheckLimits("gCoreDens1");
            AvCoreDensity();
            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = "Dataset selected: " + Name + " -- " + value });
        }
        private void AvCoreDensity()
        {
            int nCount = 0; double dSum = 0, dtmp;

            if (!(_objectsService.MfgFinishedGoods.dr["FG Core Density 1"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["FG Core Density 1"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["FG Core Density 2"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["FG Core Density 2"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["FG Core Density 3"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["FG Core Density 3"]; }

            dtmp = dSum / (double)nCount;
            if (Double.IsNaN(dtmp) || nCount == 0) { gCoreDensity = String.Empty; _objectsService.MfgFinishedGoods.dr["FG Core Density"] = DBNull.Value; }
            else { gCoreDensity = dtmp.ToString(MfgFinishedGoods.sOr); _objectsService.MfgFinishedGoods.dr["FG Core Density"] = dtmp; }

            CheckLimits("gCoreDensity");

        }
        public IActionResult OnPostGCoreDenAv2_LF(string Name, string Value)
        {
            string sFld = String.Empty;
            string value = Value;
            string sName = Name;

            switch (sName)
            {
                case "gMass2": _objectsService.MfgFinishedGoods.dr["Mass 2 FG"] = value; break;
                case "gL1_2": _objectsService.MfgFinishedGoods.dr["L1 2 FG"] = value; break;
                //            _objectsService.MfgFinishedGoods.dr[ubleField]=valuesText, "L2 2 FG"); break;
                case "gW1_2": _objectsService.MfgFinishedGoods.dr["W1 2 FG"] = value; break;
                //            _objectsService.MfgFinishedGoods.dr[ubleField]=valuesText, "W2 2 FG"); break;
                case "gT1_2": _objectsService.MfgFinishedGoods.dr["T1 2 FG"] = value; break;
                case "gT2_2": _objectsService.MfgFinishedGoods.dr["T2 2 FG"] = value; break;
                case "gT3_2": _objectsService.MfgFinishedGoods.dr["T3 2 FG"] = value; break;
                case "gT4_2": _objectsService.MfgFinishedGoods.dr["T4 2 FG"] = value; break;
                case "gT5_2": _objectsService.MfgFinishedGoods.dr["T5 2 FG"] = value; break;
            }


            int nCount = 0; double dSum = 0, dtmp;
            double dm = 0, dl, dw, dt; int nm, nl, nw, nt;

            nm = nl = nw = nt = 0;
            dm = dl = dw = dt = 0.0;

            if (!(_objectsService.MfgFinishedGoods.dr["Mass 2 FG"] == DBNull.Value)) { dm = (double)_objectsService.MfgFinishedGoods.dr["Mass 2 FG"]; nm = 1; }

            nCount = 0; dSum = 0;
            if (!(_objectsService.MfgFinishedGoods.dr["L1 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["L1 2 FG"]; }
            //            if (!(dr["L2 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)dr["L2 2 FG"]; }
            if (nCount > 0) { dl = dSum / (double)nCount; nl = 2; }

            nCount = 0; dSum = 0;
            if (!(_objectsService.MfgFinishedGoods.dr["W1 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["W1 2 FG"]; }
            //            if (!(dr["W2 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)dr["W2 2 FG"]; }
            if (nCount > 0) { dw = dSum / (double)nCount; nw = 2; }

            nCount = 0; dSum = 0;
            if (!(_objectsService.MfgFinishedGoods.dr["T1 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T1 2 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T2 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T2 2 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T3 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T3 2 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T4 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T4 2 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T5 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T5 2 FG"]; }

            if (nCount > 3) { dt = dSum / (double)nCount; nt = 5; }

            if (nm * nl * nw * nt > 0 && dl * dw * dt > 0.0) { dtmp = dm / (dl * dw * dt) * 3.809590998; gCoreDens2 = dtmp.ToString(MfgFinishedGoods.sOr); _objectsService.MfgFinishedGoods.dr["FG Core Density 2"] = dtmp; }
            else { gCoreDens2 = String.Empty; _objectsService.MfgFinishedGoods.dr["FG Core Density 2"] = DBNull.Value; }

            CheckLimits("gCoreDens2");

            AvCoreDensity();
            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new
            {
                message = "Dataset selected: " + Name + " -- " + value
            });
        }
        public IActionResult OnPostCoreDenAv3_LF(string Name, string value)
        {
            string sFld = String.Empty;
            string sText = value;
            string sName = Name;

            switch (sName)
            {
                case "gMass3": _objectsService.MfgFinishedGoods.dr["Mass 3 FG"] = value; break;
                case "gL1_3": _objectsService.MfgFinishedGoods.dr["L1 3 FG"] = value; break;
                //                case "gL2_3": _objectsServicedr[oubleFiel]=value(sText, "L2 3 FG"); break;
                case "gW1_3": _objectsService.MfgFinishedGoods.dr["W1 3 FG"] = value; break;
                //              case "gW2_3": _objectsService.Mdr[bleField(]=valueText, "W2 3 FG"); break;
                case "gT1_3": _objectsService.MfgFinishedGoods.dr["T1 3 FG"] = value; break;
                case "gT2_3": _objectsService.MfgFinishedGoods.dr["T2 3 FG"] = value; break;
                case "gT3_3": _objectsService.MfgFinishedGoods.dr["T3 3 FG"] = value; break;
                case "gT4_3": _objectsService.MfgFinishedGoods.dr["T4 3 FG"] = value; break;
                case "gT5_3": _objectsService.MfgFinishedGoods.dr["T5 3 FG"] = value; break;
            }

            int nCount = 0; double dSum = 0, dtmp;
            double dm = 0, dl, dw, dt; int nm, nl, nw, nt;

            nm = nl = nw = nt = 0;
            dm = dl = dw = dt = 0.0;

            if (!(_objectsService.MfgFinishedGoods.dr["Mass 3 FG"] == DBNull.Value)) { dm = (double)_objectsService.MfgFinishedGoods.dr["Mass 3 FG"]; nm = 1; }

            nCount = 0; dSum = 0;
            if (!(_objectsService.MfgFinishedGoods.dr["L1 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["L1 3 FG"]; }
            //            if (!(_objectsService.MfgFinishedGoods.dr["L2 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["L2 3 FG"]; }
            if (nCount > 0) { dl = dSum / (double)nCount; nl = 2; }

            nCount = 0; dSum = 0;
            if (!(_objectsService.MfgFinishedGoods.dr["W1 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["W1 3 FG"]; }
            //            if (!(_objectsService.MfgFinishedGoods.dr["W2 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["W2 3 FG"]; }
            if (nCount > 0) { dw = dSum / (double)nCount; nw = 2; }

            nCount = 0; dSum = 0;
            if (!(_objectsService.MfgFinishedGoods.dr["T1 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T1 3 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T2 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T2 3 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T3 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T3 3 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T4 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T4 3 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["T5 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["T5 3 FG"]; }

            if (nCount > 3) { dt = dSum / (double)nCount; nt = 5; }

            if (nm * nl * nw * nt > 0 && dl * dw * dt > 0.0) { dtmp = dm / (dl * dw * dt) * 3.809590998; gCoreDens3 = dtmp.ToString(MfgFinishedGoods.sOr); _objectsService.MfgFinishedGoods.dr["FG Core Density 3"] = dtmp; }
            else { gCoreDens3 = String.Empty; _objectsService.MfgFinishedGoods.dr["FG Core Density 3"] = DBNull.Value; }

            CheckLimits("gCoreDens3");

            AvCoreDensity();
            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new
            {
                message = "Dataset selected: " + Name + " -- " + value
            });
        }
        public IActionResult OnPostGThickness_LF(string Name, string value)
        {
            string sFld = String.Empty;
            string sText = value;
            string sName = Name;

            _objectsService.MfgFinishedGoods.bDataSetChanged = true;

            switch (sName)
            {
                case "gThicknessFG_1": _objectsService.MfgFinishedGoods.dr["Thickness FG - 1"] = value; break;
                case "gThicknessFG_2": _objectsService.MfgFinishedGoods.dr["Thickness FG - 2"] = value; break;
                case "gThicknessFG_3": _objectsService.MfgFinishedGoods.dr["Thickness FG - 3"] = value; break;
                case "gThicknessFG_4": _objectsService.MfgFinishedGoods.dr["Thickness FG - 4"] = value; break;
                case "gThicknessFG_5": _objectsService.MfgFinishedGoods.dr["Thickness FG - 5"] = value; break;
                case "gThicknessFG_6": _objectsService.MfgFinishedGoods.dr["Thickness FG - 6"] = value; break;
                case "gThicknessFG_7": _objectsService.MfgFinishedGoods.dr["Thickness FG - 7"] = value; break;
                case "gThicknessFG_8": _objectsService.MfgFinishedGoods.dr["Thickness FG - 8"] = value; break;
                case "gThicknessFG_9": _objectsService.MfgFinishedGoods.dr["Thickness FG - 9"] = value; break;
                case "gThicknessFG_10": _objectsService.MfgFinishedGoods.dr["Thickness FG - 10"] = value; break;
                case "gThicknessFG_11": _objectsService.MfgFinishedGoods.dr["Thickness FG - 11"] = value; break;
                case "gThicknessFG_12": _objectsService.MfgFinishedGoods.dr["Thickness FG - 12"] = value; break;
                case "gThicknessFG_13": _objectsService.MfgFinishedGoods.dr["Thickness FG - 13"] = value; break;
                case "gThicknessFG_14": _objectsService.MfgFinishedGoods.dr["Thickness FG - 14"] = value; break;
                case "gThicknessFG_15": _objectsService.MfgFinishedGoods.dr["Thickness FG - 15"] = value; break;
                case "gThicknessFG_16": _objectsService.MfgFinishedGoods.dr["Thickness FG - 16"] = value; break;
                case "gThicknessFG_17": _objectsService.MfgFinishedGoods.dr["Thickness FG - 17"] = value; break;

            }

            int nCount = 0; double dSum = 0, dtmp;
            string sField = "Thickness IP - ";
            double dmin = double.MaxValue, dmax = double.MinValue;

            for (int i = 1; i < 18; i++)
            {
                sField = "Thickness FG - " + i.ToString();
                if (!(_objectsService.MfgFinishedGoods.dr[sField] == DBNull.Value))
                {
                    dtmp = (double)_objectsService.MfgFinishedGoods.dr[sField]; dSum += dtmp; nCount += 1;
                    if (dtmp < dmin) dmin = dtmp; if (dtmp > dmax) dmax = dtmp;
                }
            }
            dtmp = dSum / (double)nCount;

            if (Double.IsNaN(dtmp) || nCount < 17)
            {
                gThickness = gFlatness = String.Empty;
                _objectsService.MfgFinishedGoods.dr["Thickness Avg FG"] = _objectsService.MfgFinishedGoods.dr["thickness valleys FG"] = _objectsService.MfgFinishedGoods.dr["thickness peaks FG"] = _objectsService.MfgFinishedGoods.dr["Flatness FG"] = DBNull.Value;
            }
            else
            {
                gThickness = dtmp.ToString(MfgFinishedGoods.sOr); _objectsService.MfgFinishedGoods.dr["Thickness Avg FG"] = dtmp;
                gFlatness = (dmin - dmax).ToString(MfgFinishedGoods.sOr); _objectsService.MfgFinishedGoods.dr["Flatness FG"] = dmin - dmax;
                _objectsService.MfgFinishedGoods.dr["thickness valleys FG"] = dmin; _objectsService.MfgFinishedGoods.dr["thickness peaks FG"] = dmax;
            }

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            CheckLimits("gThickness");
            return new JsonResult(new
            {
                message = "Dataset selected: " + Name + " -- " + value
            });
        }

        public IActionResult OnPostGQCTimesDateTime_LF(string value)
        {
            string gQCTimesDateTime = value;
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;

            if (string.IsNullOrEmpty(gQCTimesDateTime))
            {


                _objectsService.MfgFinishedGoods.dr["Next Day QC Collection Time FG"] = DBNull.Value;
            }
            else
            {
                if (DateTime.TryParseExact(gQCTimesDateTime, "yyyy-MM-ddTHH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDateTime))
                {
                    _objectsService.MfgFinishedGoods.dr["Next Day QC Collection Time FG"] = parsedDateTime.ToString();

                }
                else
                {

                    _objectsService.MfgFinishedGoods.dr["Next Day QC Collection Time FG"] = DBNull.Value;
                }
            }

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = "Success" });
        }


        public IActionResult OnPostGenInfoQC_Time_LF(DateTime value)
        {
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            bool bGetProcessData = false;
            gFBTimeStamp = value;
            if (gFBTimeStamp == null) _objectsService.MfgFinishedGoods.dr["Finished Board Time Stamp FG"] = DBNull.Value;
            else if (gFBTimeStamp.Value == null) _objectsService.MfgFinishedGoods.dr["Finished Board Time Stamp FG"] = DBNull.Value;
            else { _objectsService.MfgFinishedGoods.dr["Finished Board Time Stamp FG"] = gFBTimeStamp.Value; bGetProcessData = true; }
            _objectsService.MfgFinishedGoods.UpdateDataSet();

            if (bGetProcessData) CheckLimits("BoardTimeStamp");
            _objectsService.MfgHome.dt.Rows[_objectsService.Cbfile.iIDMfgIndex]["sFGTestTime"] = ((DateTime)_objectsService.MfgFinishedGoods.dr["Finished Board Time Stamp FG"]).ToString("MM/dd/yy - hh:mm tt");
           // _objectsService.MfgHome.SearchMfgDB.ItemsSource = _objectsService.MfgHome.dt.DefaultView;
           return new JsonResult(new{message=value });
        }

        public IActionResult OnPostGIPTimeNotLegible_LF(bool value)
        {
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;

            _objectsService.MfgFinishedGoods.dr["IP Time Stamp Not Legible"] = value;

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = value });
        }


        public IActionResult OnPostGBundleHeight_LF(string value)
        {
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            _objectsService.MfgFinishedGoods.dr["Bundle Height FG"]=value;
            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new {message=value});
        }


        public IActionResult OnPostGKFactorDates_LF(string name, DateTime value)
        {
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;

            switch (name)
            {
                case "gkFactor90Date_1":
                    _objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 1"] = value == DateTime.MinValue ? (object)DBNull.Value : value;
                    break;
                case "gkFactor90Date_2":
                    _objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 2"] = value == DateTime.MinValue ? (object)DBNull.Value : value;
                    break;
                case "gkFactor90Date_3":
                    _objectsService.MfgFinishedGoods.dr["k Factor 90 Date FG 3"] = value == DateTime.MinValue ? (object)DBNull.Value : value;
                    break;
                case "gkFactor180Date_1":
                    _objectsService.MfgFinishedGoods.dr["k Factor 180 Date FG 1"] = value == DateTime.MinValue ? (object)DBNull.Value : value;
                    break;
                case "gkFactor180Date_2":
                    _objectsService.MfgFinishedGoods.dr["k Factor 180 Date FG 2"] = value == DateTime.MinValue ? (object)DBNull.Value : value;
                    break;
                case "gkFactor180Date_3":
                    _objectsService.MfgFinishedGoods.dr["k Factor 180 Date FG 3"] = value == DateTime.MinValue ? (object)DBNull.Value : value;
                    break;
            }

            SetDateTimeField();
            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = $"{name}: {value}" });
        }

        public IActionResult OnPostGkFactor90_LF(string Name, string Value)
        {

            string sText = Value ;
            string sName = Name;

            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            switch (sName)
            {
                case "gkFactor90_1": _objectsService.MfgFinishedGoods.dr["k Factor 90 FG 1"]=Value; break;
                case "gkFactor90_2": _objectsService.MfgFinishedGoods.dr["k Factor 90 FG 2"]=Value; break;
                case "gkFactor90_3": _objectsService.MfgFinishedGoods.dr["k Factor 90 FG 3"]=Value; break;
            }

            int nCount = 0;
            double dSum = 0, dSumR = 0, dtmpR, dtmp;

            if (!(_objectsService.MfgFinishedGoods.dr["k Factor 90 FG 1"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["k Factor 90 FG 1"]; dSumR += 1.0 / (double)_objectsService.MfgFinishedGoods.dr["k Factor 90 FG 1"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["k Factor 90 FG 2"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["k Factor 90 FG 2"]; dSumR += 1.0 / (double)_objectsService.MfgFinishedGoods.dr["k Factor 90 FG 2"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["k Factor 90 FG 3"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["k Factor 90 FG 3"]; dSumR += 1.0 / (double)_objectsService.MfgFinishedGoods.dr["k Factor 90 FG 3"]; }


            if (nCount > 1)
            {
                dtmp = (dSum / (double)nCount); _objectsService.MfgFinishedGoods.dr["k Factor 90 FG"] = dtmp;
                dtmpR = (dSumR / (double)nCount); gkFactor90  = dtmpR.ToString(MfgFinishedGoods.sOr); _objectsService.MfgFinishedGoods.dr["R Value 90 - AVG FG"] = dtmpR;
            }
            else { gkFactor90  = String.Empty; _objectsService.MfgFinishedGoods.dr["k Factor 90 FG"] = _objectsService.MfgFinishedGoods.dr["R Value 90 - AVG FG"] = DBNull.Value; }

            _objectsService.MfgFinishedGoods.UpdateDataSet();

            CheckLimits("gkFactor90");
            return new JsonResult(new { message = $"{Name}: {Value}" });

        }

        public IActionResult OnPostGkFactor180_LF(string name, string value)
        {
            string sText = value;
            string sName = name;

            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            switch (sName)
            {
                case "gkFactor180_1": _objectsService.MfgFinishedGoods.dr["k Factor 180 FG 1"]=value; break;
                case "gkFactor180_2": _objectsService.MfgFinishedGoods.dr["k Factor 180 FG 2"]=value; break;
                case "gkFactor180_3": _objectsService.MfgFinishedGoods.dr["k Factor 180 FG 3"]=value; break;
            }

            int nCount = 0;
            double dSum = 0, dSumR = 0, dtmp, dtmpR;

            if (!(_objectsService.MfgFinishedGoods.dr["k Factor 180 FG 1"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["k Factor 180 FG 1"]; dSumR += 1.0 / (double)_objectsService.MfgFinishedGoods.dr["k Factor 180 FG 1"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["k Factor 180 FG 2"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["k Factor 180 FG 2"]; dSumR += 1.0 / (double)_objectsService.MfgFinishedGoods.dr["k Factor 180 FG 2"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["k Factor 180 FG 3"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["k Factor 180 FG 3"]; dSumR += 1.0 / (double)_objectsService.MfgFinishedGoods.dr["k Factor 180 FG 3"]; }


            if (nCount > 1)
            {
                dtmp = (dSum / (double)nCount); _objectsService.MfgFinishedGoods.dr["k Factor 180 FG"] = dtmp;
                dtmpR = (dSumR / (double)nCount); _objectsService.MfgFinishedGoods.dr["R Value 180 - AVG FG"] = dtmpR; gkFactor180  = dtmpR.ToString(MfgFinishedGoods.sOr);
            }
            else { gkFactor180  = String.Empty; _objectsService.MfgFinishedGoods.dr["k Factor 180 FG"] = _objectsService.MfgFinishedGoods.dr["R Value 180 - AVG FG"] = DBNull.Value; }

            _objectsService.MfgFinishedGoods.UpdateDataSet();

            CheckLimits("gkFactor180");
            return new JsonResult(new { message = $"{name}: {value}" });
        }

        public IActionResult OnPostGBoardDims_LF(string name, string value)
        {
            string sFld = String.Empty;
            string sText = value ;
            string sName = name;
            bool bDia = false;
            double dtmp;

            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            switch (sName)
            {

                case "gFGLength": _objectsService.MfgFinishedGoods.dr[ "Length FG"]=value; break;
                case "gFGWidth": _objectsService.MfgFinishedGoods.dr["Width FG"] = value; break;
                case "gFGDiagoanl1": _objectsService.MfgFinishedGoods.dr["Diagonal FG 1"]=value; bDia = true; break;
                case "gFGDiagoanl2": _objectsService.MfgFinishedGoods.dr["Diagonal FG 2"]=value; bDia = true; break;
            }

            if (!(_objectsService.MfgFinishedGoods.dr["Diagonal FG 1"] == DBNull.Value) && !(_objectsService.MfgFinishedGoods.dr["Diagonal FG 2"] == DBNull.Value))
            { dtmp = Math.Abs((double)_objectsService.MfgFinishedGoods.dr["Diagonal FG 1"] - (double)_objectsService.MfgFinishedGoods.dr["Diagonal FG 2"]); _objectsService.MfgFinishedGoods.dr["Diagonal Diff FG"] = dtmp; gFGDiagonalDiff  = dtmp.ToString(MfgFinishedGoods.sOr); }
            else { _objectsService.MfgFinishedGoods.dr["Diagonal Diff FG"] = DBNull.Value; gFGDiagonalDiff  = string.Empty; }

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = $"{name}: {value}" });
        }
        public IActionResult OnPostGBundleTemps_LF(string name, string value)
        {
            string sFld = String.Empty;
            string sText = value;
            string sName = name;

            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            switch (sName)
            {
                case "gLoggerID": _objectsService.MfgFinishedGoods.dr["Logger ID # FG"]=value; break;
                case "gInitProbeTemp": _objectsService.MfgFinishedGoods.dr["Initial Probe Temp FG"]=value; break;
                case "gMaxProbeTemp": _objectsService.MfgFinishedGoods.dr["Max Probe Temp FG"]=value; break;
                case "gFinalProbeTemp": _objectsService.MfgFinishedGoods.dr["Final Probe Temp FG"]=value; break;
            }

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = $"{name}: {value}" });
        }

        public IActionResult OnPostGNailPull_LF(string name, string value)
        {
            string sFld = String.Empty;
            string sText = value ;
            string sName = name;

            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            switch (sName)
            {
                case "gNailPull_1": _objectsService.MfgFinishedGoods.dr["Nail Pull FG 1"]=value; break;
                case "gNailPull_2": _objectsService.MfgFinishedGoods.dr["Nail Pull FG 2"]=value; break;
                case "gNailPull_3": _objectsService.MfgFinishedGoods.dr["Nail Pull FG 3"]=value; break;
            }
            int nCount = 0;
            double dSum = 0, dtmp;
            if (!(_objectsService.MfgFinishedGoods.dr["Nail Pull FG 1"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["Nail Pull FG 1"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["Nail Pull FG 2"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["Nail Pull FG 2"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["Nail Pull FG 3"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["Nail Pull FG 3"]; }

            if (nCount == 3) { dtmp = (dSum / (double)nCount); gNailPull  = dtmp.ToString(MfgFinishedGoods.sOr); _objectsService.MfgFinishedGoods.dr["Nail Pull FG"] = dtmp; }
            else { gNailPull  = String.Empty; _objectsService.MfgFinishedGoods.dr["Nail Pull FG"] = DBNull.Value; }


            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = $"{name}: {value}" });
        }

        public IActionResult OnPostGFacerPeel_LF(string name, string value)
        {

            string sFld = String.Empty;
            string sText = value;
            string sName = name;

            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            switch (sName)
            {
                case "gFacerPeel1": _objectsService.MfgFinishedGoods.dr["Facer Peel 1 FG"]=value; break;
                case "gFacerPeel2": _objectsService.MfgFinishedGoods.dr["Facer Peel 2 FG"]=value; break;
                case "gFacerPeel3": _objectsService.MfgFinishedGoods.dr["Facer Peel 3 FG"]=value; break;

            }

            int nCount = 0;
            double dSum = 0, dtmp;
            if (!(_objectsService.MfgFinishedGoods.dr["Facer Peel 1 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["Facer Peel 1 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["Facer Peel 2 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["Facer Peel 2 FG"]; }
            if (!(_objectsService.MfgFinishedGoods.dr["Facer Peel 3 FG"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgFinishedGoods.dr["Facer Peel 3 FG"]; }

            if (nCount == 3) { dtmp = (dSum / (double)nCount); gFacerPeelAvg  = gFacerPeelAvg_QC  = dtmp.ToString(MfgFinishedGoods.sOr); _objectsService.MfgFinishedGoods.dr["Facer Peel FG"] = dtmp; }
            else { gFacerPeelAvg  = gFacerPeelAvg_QC  = String.Empty; _objectsService.MfgFinishedGoods.dr["Facer Peel FG"] = DBNull.Value; }

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            CheckLimits("gFacerPeelAvg");
            return new JsonResult(new { message = $"{name}: {value}" });
        }


        #region WindowForms DateTime functions

        public IActionResult OnPostGkFactorTime1_LF(DateTime value)
        {
            gkFactorTime1 = value;
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;

            if (gkFactorTime1 == null) _objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 1"] = DBNull.Value;
            else if (gkFactorTime1.Value == null) _objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 1"] = DBNull.Value;
            else _objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 1"] = gkFactorTime1.Value;

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = value });
        }

        public IActionResult OnPostGkFactorTime2_LF(DateTime value)
        {
            gkFactorTime2 = value;
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;

            if (gkFactorTime2 == null) _objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 2"] = DBNull.Value;
            else if (gkFactorTime2.Value == null) _objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 2"] = DBNull.Value;
            else _objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 2"] = gkFactorTime2.Value;

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = value });
        }

        public IActionResult OnPostGkFactorTime3_LF(DateTime value)
        {
            gkFactorTime3 = value;
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;

            if (gkFactorTime3 == null) _objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 3"] = DBNull.Value;
            else if (gkFactorTime3.Value == null) _objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 3"] = DBNull.Value;
            else _objectsService.MfgFinishedGoods.dr["k Factor DateTime FG 3"] = gkFactorTime3.Value;

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = value });
        }



        private void gTime_GotFocus(object sender, EventArgs e)
        {
            //  System.Windows.Forms.DateTimePicker dtp = sender as System.Windows.Forms.DateTimePicker;
            // if (dtp.CustomFormat == sNull) { dtp.CustomFormat = sTimeFormat; dtp.Value = DateTime.Now; }
        }

        public DateTime gDateTime_GF()
        {
            //System.Windows.Forms.DateTimePicker dtp = sender as System.Windows.Forms.DateTimePicker;
            //if (_objectsService.dtp.CustomFormat == sNull) { dtp.CustomFormat = sDateTimeFormat; dtp.Value = DateTime.Now; }
            return DateTime.Now;
        }



        public IActionResult OnPostGRetestQCTime_LF(DateTime value)
        {
            gRetestQCTime = value;
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            if (gRetestQCTime == null) _objectsService.MfgFinishedGoods.dr["Retest QC Collection Time FG"] = DBNull.Value;
            else if (gRetestQCTime.Value == null) _objectsService.MfgFinishedGoods.dr["Retest QC Collection Time FG"] = DBNull.Value;
            else _objectsService.MfgFinishedGoods.dr["Retest QC Collection Time FG"] = gRetestQCTime.Value;

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = value });
        }

        public IActionResult OnPostGInitProbeTime_LF(DateTime value)
        {
            gInitProbeTime = value;
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;

            if (gInitProbeTime == null) _objectsService.MfgFinishedGoods.dr["Initial Probe Time FG"] = DBNull.Value;
            else if (gInitProbeTime.Value == null) _objectsService.MfgFinishedGoods.dr["Initial Probe Time FG"] = DBNull.Value;
            else _objectsService.MfgFinishedGoods.dr["Initial Probe Time FG"] = gInitProbeTime.Value;

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = value });
        }

        public IActionResult OnPostGMaxTempTimeInit_LF(DateTime value)
        {
            gMaxTempTimeInit = value;
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            if (gMaxTempTimeInit == null) _objectsService.MfgFinishedGoods.dr["Max Probe Time - Initial FG"] = DBNull.Value;
            else if (gMaxTempTimeInit.Value == null) _objectsService.MfgFinishedGoods.dr["Max Probe Time - Initial FG"] = DBNull.Value;
            else _objectsService.MfgFinishedGoods.dr["Max Probe Time - Initial FG"] = gMaxTempTimeInit.Value;

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = value });
        }

        public IActionResult OnPostGMaxTempTimeFinal_LF(DateTime value)
        {
            gMaxTempTimeFinal = value;
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;

            if (gMaxTempTimeFinal == null) _objectsService.MfgFinishedGoods.dr["Max Probe Time - Final FG"] = DBNull.Value;
            else if (gMaxTempTimeFinal.Value == null) _objectsService.MfgFinishedGoods.dr["Max Probe Time - Final FG"] = DBNull.Value;
            else _objectsService.MfgFinishedGoods.dr["Max Probe Time - Final FG"] = gMaxTempTimeFinal.Value;

            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = value });
        }
        public IActionResult OnPostGFinalProbeTime_LF(DateTime value)
        {
            gFinalProbeTime = value;
            _objectsService.MfgFinishedGoods.bDataSetChanged = true;
            if (gFinalProbeTime == null) _objectsService.MfgFinishedGoods.dr["Final Probe Time FG"] = DBNull.Value;
            else if (gFinalProbeTime.Value == null) _objectsService.MfgFinishedGoods.dr["Final Probe Time FG"] = DBNull.Value;
            else _objectsService.MfgFinishedGoods.dr["Final Probe Time FG"] = gFinalProbeTime.Value;
            _objectsService.MfgFinishedGoods.UpdateDataSet();
            return new JsonResult(new { message = value });
        }




        #endregion














        public string SaveData()
        {
            string sMsg;


            return gThicknessFG_1 + " - " + DateTime.Now.Second.ToString();

        }
    }
}