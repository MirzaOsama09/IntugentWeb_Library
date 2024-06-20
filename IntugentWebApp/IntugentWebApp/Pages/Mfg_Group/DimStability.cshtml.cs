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
    public class DimensionStabilityModel : PageModel
    {
        public string gID { get; set; }
        public bool gDataSetNextIsEnabled { get; set; }
        public bool gDataSetPrevIsEnabled { get; set; }
        public bool gDimStabilityDoneIsChecked { get; set; }
        public string gProductionDate { get; set; }
        public string gProductionTime { get; set; }
        public string gProductCode { get; set; }
        public bool gEdgeH1IsChecked { get; set; }
        public bool gEdgeH2IsChecked { get; set; }
        public bool gEdgeC1IsChecked { get; set; }
        public bool gEdgeC2IsChecked { get; set; }

        public string gInit_H1_L1 { get; set; }
        public string gInit_H1_L2 { get; set; }
        public string gInit_H1_L3 { get; set; }
        public string gInit_H1_W1 { get; set; }
        public string gInit_H1_W2 { get; set; }
        public string gInit_H1_W3 { get; set; }
        public string gInit_H1_T1 { get; set; }
        public string gInit_H1_T2 { get; set; }
        public string gInit_H1_T3 { get; set; }
        public string gInit_H1_T4 { get; set; }
        public string gInit_H1_T5 { get; set; }
        public string gFinal_H1_L1 { get; set; }
        public string gFinal_H1_L2 { get; set; }
        public string gFinal_H1_L3 { get; set; }
        public string gFinal_H1_W1 { get; set; }
        public string gFinal_H1_W2 { get; set; }
        public string gFinal_H1_W3 { get; set; }
        public string gFinal_H1_T1 { get; set; }
        public string gFinal_H1_T2 { get; set; }
        public string gFinal_H1_T3 { get; set; }
        public string gFinal_H1_T4 { get; set; }
        public string gFinal_H1_T5 { get; set; }

        public string gInit_H2_L1 { get; set; }
        public string gInit_H2_L2 { get; set; }
        public string gInit_H2_L3 { get; set; }
        public string gInit_H2_W1 { get; set; }
        public string gInit_H2_W2 { get; set; }
        public string gInit_H2_W3 { get; set; }
        public string gInit_H2_T1 { get; set; }
        public string gInit_H2_T2 { get; set; }
        public string gInit_H2_T3 { get; set; }
        public string gInit_H2_T4 { get; set; }
        public string gInit_H2_T5 { get; set; }
        public string gFinal_H2_L1 { get; set; }
        public string gFinal_H2_L2 { get; set; }
        public string gFinal_H2_L3 { get; set; }
        public string gFinal_H2_W1 { get; set; }
        public string gFinal_H2_W2 { get; set; }
        public string gFinal_H2_W3 { get; set; }
        public string gFinal_H2_T1 { get; set; }
        public string gFinal_H2_T2 { get; set; }
        public string gFinal_H2_T3 { get; set; }
        public string gFinal_H2_T4 { get; set; }
        public string gFinal_H2_T5 { get; set; }
        public string gInit_C1_L1 { get; set; }
        public string gInit_C1_L2 { get; set; }
        public string gInit_C1_L3 { get; set; }
        public string gInit_C1_W1 { get; set; }
        public string gInit_C1_W2 { get; set; }
        public string gInit_C1_W3 { get; set; }
        public string gInit_C1_T1 { get; set; }
        public string gInit_C1_T2 { get; set; }
        public string gInit_C1_T3 { get; set; }
        public string gInit_C1_T4 { get; set; }
        public string gInit_C1_T5 { get; set; }
        public string gFinal_C1_L1 { get; set; }
        public string gFinal_C1_L2 { get; set; }
        public string gFinal_C1_L3 { get; set; }
        public string gFinal_C1_W1 { get; set; }
        public string gFinal_C1_W2 { get; set; }
        public string gFinal_C1_W3 { get; set; }
        public string gFinal_C1_T1 { get; set; }
        public string gFinal_C1_T2 { get; set; }
        public string gFinal_C1_T3 { get; set; }
        public string gFinal_C1_T4 { get; set; }
        public string gFinal_C1_T5 { get; set; }
        public string gInit_C2_L1 { get; set; }
        public string gInit_C2_L2 { get; set; }
        public string gInit_C2_L3 { get; set; }
        public string gInit_C2_W1 { get; set; }
        public string gInit_C2_W2 { get; set; }
        public string gInit_C2_W3 { get; set; }
        public string gInit_C2_T1 { get; set; }
        public string gInit_C2_T2 { get; set; }
        public string gInit_C2_T3 { get; set; }
        public string gInit_C2_T4 { get; set; }
        public string gInit_C2_T5 { get; set; }
        public string gFinal_C2_L1 { get; set; }
        public string gFinal_C2_L2 { get; set; }
        public string gFinal_C2_L3 { get; set; }
        public string gFinal_C2_W1 { get; set; }
        public string gFinal_C2_W2 { get; set; }
        public string gFinal_C2_W3 { get; set; }
        public string gFinal_C2_T1 { get; set; }
        public string gFinal_C2_T2 { get; set; }
        public string gFinal_C2_T3 { get; set; }
        public string gFinal_C2_T4 { get; set; }
        public string gFinal_C2_T5 { get; set; }
        public string gInit_WF_L1 { get; set; }
        public string gInit_WF_L2 { get; set; }
        public string gInit_WF_L3 { get; set; }
        public string gInit_WF_W1 { get; set; }
        public string gInit_WF_W2 { get; set; }
        public string gInit_WF_W3 { get; set; }
        public string gFinal_WF_L1 { get; set; }
        public string gFinal_WF_L2 { get; set; }
        public string gFinal_WF_L3 { get; set; }
        public string gFinal_WF_W1 { get; set; }
        public string gFinal_WF_W2 { get; set; }
        public string gFinal_WF_W3 { get; set; }
        public string gSide1_Depth1 { get; set; }
        public string gSide1_Depth2 { get; set; }
        public string gSide1_Depth3 { get; set; }
        public string gSide1_Depth4 { get; set; }
        public string gSide1_Depth5 { get; set; }
        public string gSide2_Depth1 { get; set; }
        public string gSide2_Depth2 { get; set; }
        public string gSide2_Depth3 { get; set; }
        public string gSide2_Depth4 { get; set; }
        public string gSide2_Depth5 { get; set; }
        public string gChangeOvenLength { get; set; }
        public string gChangeOvenWidth { get; set; }
        public string gChangeOvenThickness { get; set; }
        public string gAvgDepth { get; set; }
        public string gMaxDepth { get; set; }
        public string gChangeWFLength { get; set; }
        public string gChangeWFWidth { get; set; }
        public string gDeviation { get; set; }
        public string gChangeFreezerLength { get; set; }
        public string gChangeFreezerWidth { get; set; }
        public string gChangeFreezerThickness { get; set; }
        public DateTime? gTestDateTime { get; set; }
        public string gDevType { get; set; }

        private readonly ObjectsService _objectsService;
        public DimensionStabilityModel(ObjectsService objectsService)
        {
            _objectsService = objectsService;
        }
        public void OnGet()
        {
            _objectsService.MfgDimensionsStability.GetDataSet();
            SetView();
        }

        public void SetView()
        {
            if (_objectsService.Cbfile.iIDMfgIndex == 0) gDataSetNextIsEnabled = false; else gDataSetNextIsEnabled = true;
            if (_objectsService.Cbfile.iIDMfgIndex == _objectsService.MfgHome.dt.Rows.Count - 1) gDataSetPrevIsEnabled = false; else gDataSetPrevIsEnabled = true;

            //            GetDataSet();
            //            CPages.PageInProcess_1.GetDataSet();
                       _objectsService.MfgDimensionsStability.drIP = _objectsService.MfgInProcess.dr;
            //           CPages.PageFinishedGoods_1.GetDataSet();
                        _objectsService.MfgDimensionsStability.drFG = _objectsService.MfgFinishedGoods.dr;

            #region gen info, edge collapse, % change

            gID = _objectsService.MfgDimensionsStability.dr["ID4ALL"].ToString();
            if ((_objectsService.MfgDimensionsStability.dr["DS Testing Complete"] == DBNull.Value)) gDimStabilityDoneIsChecked = false; else gDimStabilityDoneIsChecked = (bool)_objectsService.MfgDimensionsStability.dr["DS Testing Complete"];
            if (_objectsService.MfgDimensionsStability.drIP?["Test Date"] == DBNull.Value) gProductionDate = String.Empty; else gProductionDate = ((DateTime)_objectsService.MfgDimensionsStability.drIP["Test Date"]).ToString("yyyy-MM-ddTHH:mm");
            if (_objectsService.MfgDimensionsStability.drFG?["Finished Board Time Stamp FG"] == DBNull.Value) gProductionTime = String.Empty; else gProductionTime = ((DateTime)_objectsService.MfgDimensionsStability.drFG["Finished Board Time Stamp FG"]).ToString("yyyy-MM-ddTHH:mm");

            //           if (_objectsService.MfgDimensionsStability.drIP["Test Time"] == DBNull.Value) gProductionTime.Text = String.Empty; else gProductionTime.Text = ((DateTime)_objectsService.MfgDimensionsStability.drIP["Test Time"]).ToString("hh:mm tt");
            if (_objectsService.MfgDimensionsStability.drIP["Product ID"] == DBNull.Value) gProductCode = String.Empty; else gProductCode = _objectsService.MfgDimensionsStability.drIP["Product ID"].ToString();

            if (_objectsService.MfgDimensionsStability.dr["Edge H1"] == DBNull.Value) gEdgeH1IsChecked = false; else gEdgeH1IsChecked = (bool)_objectsService.MfgDimensionsStability.dr["Edge H1"];
            if (_objectsService.MfgDimensionsStability.dr["Edge H2"] == DBNull.Value) gEdgeH2IsChecked = false; else gEdgeH2IsChecked = (bool)_objectsService.MfgDimensionsStability.dr["Edge H2"];
            if (_objectsService.MfgDimensionsStability.dr["Edge C1"] == DBNull.Value) gEdgeC1IsChecked = false; else gEdgeC1IsChecked = (bool)_objectsService.MfgDimensionsStability.dr["Edge C1"];
            if (_objectsService.MfgDimensionsStability.dr["Edge C2"] == DBNull.Value) gEdgeC2IsChecked = false; else gEdgeC2IsChecked = (bool)_objectsService.MfgDimensionsStability.dr["Edge C2"];

            gChangeOvenLength = _objectsService.MfgDimensionsStability.SetDoubleTextField("% Change Length Oven", MfgDimStability.sOr);
            gChangeOvenWidth = _objectsService.MfgDimensionsStability.SetDoubleTextField("% Change Width Oven", MfgDimStability.sOr);
            gChangeOvenThickness = _objectsService.MfgDimensionsStability.SetDoubleTextField("% Change Thickness Oven", MfgDimStability.sOr);

            gChangeFreezerLength = _objectsService.MfgDimensionsStability.SetDoubleTextField("% Change Length Freezer", MfgDimStability.sOr);
            gChangeFreezerWidth = _objectsService.MfgDimensionsStability.SetDoubleTextField("% Change Width Freezer", MfgDimStability.sOr);
            gChangeFreezerThickness = _objectsService.MfgDimensionsStability.SetDoubleTextField("% Change Thickness Freezer", MfgDimStability.sOr);
            gAvgDepth = _objectsService.MfgDimensionsStability.SetDoubleTextField("AvgDepth", MfgDimStability.sOr);
            gMaxDepth = _objectsService.MfgDimensionsStability.SetDoubleTextField("MaxDepth", MfgDimStability.sOr);
            gChangeWFLength = _objectsService.MfgDimensionsStability.SetDoubleTextField("%ChangeLengthWF", MfgDimStability.sOr);
            gChangeWFWidth = _objectsService.MfgDimensionsStability.SetDoubleTextField("%ChangeWidthWF", MfgDimStability.sOr);

            if (_objectsService.MfgDimensionsStability.dr["DateSampleIn"] == DBNull.Value) gTestDateTime = null;
            else { gTestDateTime = (DateTime)_objectsService.MfgDimensionsStability.dr["DateSampleIn"]; }
            gDeviation = _objectsService.MfgDimensionsStability.SetDoubleTextField("DevFromTable");
            if (_objectsService.MfgDimensionsStability.dr["DevType"] == DBNull.Value) gDevType = null; else gDevType = (string)_objectsService.MfgDimensionsStability.dr["DevType"];




            #endregion

            #region Heater 1

            gInit_H1_L1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - L1");
            gInit_H1_L2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - L2");
            gInit_H1_L3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - L3");

            gInit_H1_W1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - W1");
            gInit_H1_W2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - W2");
            gInit_H1_W3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - W3");

            gInit_H1_T1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - T1");
            gInit_H1_T2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - T2");
            gInit_H1_T3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - T3");
            gInit_H1_T4 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - T4");
            gInit_H1_T5 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H1 - T5");


            gFinal_H1_L1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - L1");
            gFinal_H1_L2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - L2");
            gFinal_H1_L3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - L3");

            gFinal_H1_W1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - W1");
            gFinal_H1_W2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - W2");
            gFinal_H1_W3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - W3");

            gFinal_H1_T1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - T1");
            gFinal_H1_T2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - T2");
            gFinal_H1_T3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - T3");
            gFinal_H1_T4 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - T4");
            gFinal_H1_T5 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H1 - T5");

            #endregion

            #region Heater 2

            gInit_H2_L1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - L1");
            gInit_H2_L2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - L2");
            gInit_H2_L3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - L3");

            gInit_H2_W1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - W1");
            gInit_H2_W2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - W2");
            gInit_H2_W3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - W3");

            gInit_H2_T1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - T1");
            gInit_H2_T2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - T2");
            gInit_H2_T3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - T3");
            gInit_H2_T4 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - T4");
            gInit_H2_T5 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial H2 - T5");


            gFinal_H2_L1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - L1");
            gFinal_H2_L2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - L2");
            gFinal_H2_L3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - L3");

            gFinal_H2_W1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - W1");
            gFinal_H2_W2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - W2");
            gFinal_H2_W3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - W3");

            gFinal_H2_T1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - T1");
            gFinal_H2_T2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - T2");
            gFinal_H2_T3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - T3");
            gFinal_H2_T4 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - T4");
            gFinal_H2_T5 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final H2 - T5");

            # endregion


            #region Freezer 1

            gInit_C1_L1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - L1");
            gInit_C1_L2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - L2");
            gInit_C1_L3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - L3");

            gInit_C1_W1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - W1");
            gInit_C1_W2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - W2");
            gInit_C1_W3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - W3");

            gInit_C1_T1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - T1");
            gInit_C1_T2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - T2");
            gInit_C1_T3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - T3");
            gInit_C1_T4 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - T4");
            gInit_C1_T5 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C1 - T5");


            gFinal_C1_L1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - L1");
            gFinal_C1_L2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - L2");
            gFinal_C1_L3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - L3");

            gFinal_C1_W1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - W1");
            gFinal_C1_W2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - W2");
            gFinal_C1_W3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - W3");

            gFinal_C1_T1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - T1");
            gFinal_C1_T2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - T2");
            gFinal_C1_T3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - T3");
            gFinal_C1_T4 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - T4");
            gFinal_C1_T5 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C1 - T5");

            # endregion


            #region Heater 2

            gInit_C2_L1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - L1");
            gInit_C2_L2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - L2");
            gInit_C2_L3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - L3");

            gInit_C2_W1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - W1");
            gInit_C2_W2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - W2");
            gInit_C2_W3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - W3");

            gInit_C2_T1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - T1");
            gInit_C2_T2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - T2");
            gInit_C2_T3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - T3");
            gInit_C2_T4 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - T4");
            gInit_C2_T5 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Initial C2 - T5");


            gFinal_C2_L1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - L1");
            gFinal_C2_L2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - L2");
            gFinal_C2_L3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - L3");
            gFinal_C2_W1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - W1");
            gFinal_C2_W2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - W2");
            gFinal_C2_W3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - W3");
            gFinal_C2_T1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - T1");
            gFinal_C2_T2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - T2");
            gFinal_C2_T3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - T3");
            gFinal_C2_T4 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - T4");
            gFinal_C2_T5 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Final C2 - T5");

            #endregion

            #region WF Dimension
            gInit_WF_L1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("InitialWF-L1");
            gInit_WF_L2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("InitialWF-L2");
            gInit_WF_L3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("InitialWF-L3");
            gInit_WF_W1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("InitialWF-W1");
            gInit_WF_W2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("InitialWF-W2");
            gInit_WF_W3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("InitialWF-W3");
            gFinal_WF_L1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("FinalWF-L1");
            gFinal_WF_L2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("FinalWF-L2");
            gFinal_WF_L3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("FinalWF-L3");
            gFinal_WF_W1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("FinalWF-W1");
            gFinal_WF_W2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("FinalWF-W2");
            gFinal_WF_W3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("FinalWF-W3");

            #endregion

            #region WF Depths
            gSide1_Depth1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Loc1Depth1");
            gSide1_Depth2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Loc1Depth2");
            gSide1_Depth3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Loc1Depth3");
            gSide1_Depth4 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Loc1Depth4");
            gSide1_Depth5 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Loc1Depth5");
            gSide2_Depth1 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Loc2Depth1");
            gSide2_Depth2 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Loc2Depth2");
            gSide2_Depth3 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Loc2Depth3");
            gSide2_Depth4 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Loc2Depth4");
            gSide2_Depth5 = _objectsService.MfgDimensionsStability.SetDoubleTextField("Loc2Depth5");


            #endregion
            CheckLimits("All");

        }

        private void CheckLimits(string sF)
        {
            //Must be included in setview and   change products

            //if (sF == "All") CProdTargets.GetProductTargets();


            //if (sF == "gChangeOvenLength" || sF == "All")
            //{
            //    if (_objectsService.MfgDimensionsStability.dr["% Change Length Oven"] == DBNull.Value) gChangeOvenLength.Background = backColorCal;
            //    else if (CProdTargets.DimStabLengthWithinLimits((double)_objectsService.MfgDimensionsStability.dr["% Change Length Oven"], "hot") == "N") gChangeOvenLength.Background = backColorWarn; else gChangeOvenLength.Background = backColorCal;
            //}

            //if (sF == "gChangeOvenWidth" || sF == "All")
            //{
            //    if (_objectsService.MfgDimensionsStability.dr["% Change Width Oven"] == DBNull.Value) gChangeOvenWidth.Background = backColorCal;
            //    else if (CProdTargets.DimStabWidthWithinLimits((double)_objectsService.MfgDimensionsStability.dr["% Change Width Oven"], "hot") == "N") gChangeOvenWidth.Background = backColorWarn; else gChangeOvenWidth.Background = backColorCal;
            //}

            //if (sF == "gChangeOvenThickness" || sF == "All")
            //{
            //    if (_objectsService.MfgDimensionsStability.dr["% Change Thickness Oven"] == DBNull.Value) gChangeOvenThickness.Background = backColorCal;
            //    else if (CProdTargets.DimStabThicknessWithinLimits((double)_objectsService.MfgDimensionsStability.dr["% Change Thickness Oven"], "hot") == "N") gChangeOvenThickness.Background = backColorWarn; else gChangeOvenThickness.Background = backColorCal;
            //}

            //if (sF == "gChangeFreezerLength" || sF == "All")
            //{
            //    if (_objectsService.MfgDimensionsStability.dr["% Change Length Freezer"] == DBNull.Value) gChangeFreezerLength.Background = backColorCal;
            //    else if (CProdTargets.DimStabLengthWithinLimits((double)_objectsService.MfgDimensionsStability.dr["% Change Length Freezer"], "cold") == "N") gChangeFreezerLength.Background = backColorWarn; else gChangeFreezerLength.Background = backColorCal;
            //}

            //if (sF == "gChangeFreezerWidth" || sF == "All")
            //{
            //    if (_objectsService.MfgDimensionsStability.dr["% Change Width Freezer"] == DBNull.Value) gChangeFreezerWidth.Background = backColorCal;
            //    else if (CProdTargets.DimStabWidthWithinLimits((double)_objectsService.MfgDimensionsStability.dr["% Change Width Freezer"], "cold") == "N") gChangeFreezerWidth.Background = backColorWarn; else gChangeFreezerWidth.Background = backColorCal;
            //}

            //if (sF == "gChangeFreezerThickness" || sF == "All")
            //{
            //    if (_objectsService.MfgDimensionsStability.dr["% Change Thickness Freezer"] == DBNull.Value) gChangeFreezerThickness.Background = backColorCal;
            //    else if (CProdTargets.DimStabThicknessWithinLimits((double)_objectsService.MfgDimensionsStability.dr["% Change Thickness Freezer"], "cold") == "N") gChangeFreezerThickness.Background = backColorWarn; else gChangeFreezerThickness.Background = backColorCal;
            //}
        }

        public IActionResult OnPostGCheckBoxes_LF(string Name, string value)
        {
            string sFld = String.Empty;
            string sName = Name;

            _objectsService.MfgDimensionsStability.bDataSetChanged = true;

            switch (sName)
            {
                case "gEdgeH1":  _objectsService.MfgDimensionsStability.dr["Edge H1"] =value ; break;
                case "gEdgeH2":  _objectsService.MfgDimensionsStability.dr["Edge H2"] =value; break;
                case "gEdgeC1": _objectsService.MfgDimensionsStability.dr["Edge C1"] = value; break;
                case "gEdgeC2": _objectsService.MfgDimensionsStability.dr["Edge C2"] = value; break;
                case "gDimStabilityDone": if (_objectsService.MfgDimensionsStability.dr["DS Testing Complete"] == DBNull.Value) gDimStabilityDoneIsChecked = false; else _objectsService.MfgDimensionsStability.dr["DS Testing Complete"] = value; break;
            }
            _objectsService.MfgDimensionsStability.UpdateDataSet();
         
            return new JsonResult(new { message = Name, value , gEdgeH1IsChecked });
        }

        public IActionResult OnPostGOven1_LF(string Name, string value)
        {
            string sName = Name;
            bool bl = false, bw = false, bt = false;
            int nCount = 0;
            double dSum = 0, dSumInit = 0, dSumFinal = 0;

            _objectsService.MfgDimensionsStability.bDataSetChanged = true;
            switch (sName)
            {
                case "gInit_H1_L1": _objectsService.MfgDimensionsStability.dr["Initial H1 - L1"] = value; bl = true; break;
                case "gInit_H1_L2": _objectsService.MfgDimensionsStability.dr["Initial H1 - L2"] = value; bl = true; break;
                case "gInit_H1_L3": _objectsService.MfgDimensionsStability.dr["Initial H1 - L3"] = value; bl = true; break;
                case "gFinal_H1_L1": _objectsService.MfgDimensionsStability.dr["Final H1 - L1"] = value; bl = true; break;
                case "gFinal_H1_L2": _objectsService.MfgDimensionsStability.dr["Final H1 - L2"] = value; bl = true; break;
                case "gFinal_H1_L3": _objectsService.MfgDimensionsStability.dr["Final H1 - L3"] = value; bl = true; break;

                case "gInit_H1_W1": _objectsService.MfgDimensionsStability.dr["Initial H1 - W1"]=value; bw = true; break;
                case "gInit_H1_W2": _objectsService.MfgDimensionsStability.dr["Initial H1 - W2"]=value; bw = true; break;
                case "gInit_H1_W3": _objectsService.MfgDimensionsStability.dr["Initial H1 - W3"]=value; bw = true; break;
                case "gFinal_H1_W1": _objectsService.MfgDimensionsStability.dr["Final H1 - W1"]=value; bw = true; break;
                case "gFinal_H1_W2": _objectsService.MfgDimensionsStability.dr["Final H1 - W2"]=value; bw = true; break;
                case "gFinal_H1_W3": _objectsService.MfgDimensionsStability.dr["Final H1 - W3"]=value; bw = true; break;

                case "gInit_H1_T1": _objectsService.MfgDimensionsStability.dr["Initial H1 - T1"] = value; bt = true; break;
                case "gInit_H1_T2": _objectsService.MfgDimensionsStability.dr["Initial H1 - T2"] = value; bt = true; break;
                case "gInit_H1_T3": _objectsService.MfgDimensionsStability.dr["Initial H1 - T3"] = value; bt = true; break;
                case "gInit_H1_T4": _objectsService.MfgDimensionsStability.dr["Initial H1 - T4"] = value; bt = true; break;
                case "gInit_H1_T5": _objectsService.MfgDimensionsStability.dr["Initial H1 - T5"] = value; bt = true; break;
                case "gFinal_H1_T1": _objectsService.MfgDimensionsStability.dr["Final H1 - T1"] = value; bt = true; break;
                case "gFinal_H1_T2": _objectsService.MfgDimensionsStability.dr["Final H1 - T2"] = value; bt = true; break;
                case "gFinal_H1_T3": _objectsService.MfgDimensionsStability.dr["Final H1 - T3"] = value; bt = true; break;
                case "gFinal_H1_T4": _objectsService.MfgDimensionsStability.dr["Final H1 - T4"] = value; bt = true; break;
                case "gFinal_H1_T5": _objectsService.MfgDimensionsStability.dr["Final H1 - T5"] = value; bt = true; break;
            }
            if (bl)
            {
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - L1"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - L1"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - L1"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - L1"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - L2"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - L2"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - L2"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - L2"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - L3"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - L3"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - L3"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - L3"]; }
                if (nCount == 3 && dSumInit > 0) { _objectsService.MfgDimensionsStability.dr["%change L - H1"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else _objectsService.MfgDimensionsStability.dr["%change L - H1"] = DBNull.Value;
                DimensionChanges("Oven-L");
            }
            else if (bw)
            {
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - W1"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - W1"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - W1"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - W1"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - W2"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - W2"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - W2"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - W2"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - W3"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - W3"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - W3"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - W3"]; }
                if (nCount == 3 && dSumInit > 0) { _objectsService.MfgDimensionsStability.dr["%change W - H1"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else _objectsService.MfgDimensionsStability.dr["%change W - H1"] = DBNull.Value;
                DimensionChanges("Oven-W");
            }
            else if (bt)
            {
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - T1"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - T1"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - T1"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - T1"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - T2"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - T2"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - T2"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - T2"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - T3"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - T3"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - T3"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - T3"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - T4"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - T4"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - T4"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - T4"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial H1 - T5"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final H1 - T5"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial H1 - T5"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final H1 - T5"]; }


                if (nCount == 5 && dSumInit > 0) { _objectsService.MfgDimensionsStability.dr["%change T - H1"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else _objectsService.MfgDimensionsStability.dr["%change T - H1"] = DBNull.Value;
                DimensionChanges("Oven-T");
            }


            _objectsService.MfgDimensionsStability.UpdateDataSet();
            return new JsonResult(new { message = Name, value });
        }
        public void DimensionChanges(string sType)
        {
            int nCount = 0; double dSum = 0, dtmp;

            _objectsService.MfgDimensionsStability.bDataSetChanged = true;
            switch (sType)
            {
                case "Oven-L":
                    if (!(_objectsService.MfgDimensionsStability.dr["%change L - H1"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change L - H1"]; }
                    if (!(_objectsService.MfgDimensionsStability.dr["%change L - H2"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change L - H2"]; }
                    if (nCount == 2) { dtmp = dSum / (double)nCount; _objectsService.MfgDimensionsStability.dr["% Change Length Oven"] = dtmp; gChangeOvenLength  = dtmp.ToString("0.000"); } else { _objectsService.MfgDimensionsStability.dr["% Change Length Oven"] = DBNull.Value; gChangeOvenLength  = string.Empty; }
                    CheckLimits("gChangeOvenLength");

                    break;
                case "Oven-W":
                    if (!(_objectsService.MfgDimensionsStability.dr["%change W - H1"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change W - H1"]; }
                    if (!(_objectsService.MfgDimensionsStability.dr["%change W - H2"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change W - H2"]; }
                    if (nCount == 2) { dtmp = dSum / (double)nCount; _objectsService.MfgDimensionsStability.dr["% Change Width Oven"] = dtmp; gChangeOvenWidth  = dtmp.ToString("0.000"); } else { _objectsService.MfgDimensionsStability.dr["% Change Width Oven"] = DBNull.Value; gChangeOvenWidth  = string.Empty; }
                    CheckLimits("gChangeOvenWidth");
                    break;
                case "Oven-T":
                    if (!(_objectsService.MfgDimensionsStability.dr["%change T - H1"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change T - H1"]; }
                    if (!(_objectsService.MfgDimensionsStability.dr["%change T - H2"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change T - H2"]; }
                    if (nCount == 2) { dtmp = dSum / (double)nCount; _objectsService.MfgDimensionsStability.dr["% Change Thickness Oven"] = dtmp; gChangeOvenThickness  = dtmp.ToString("0.000"); } else { _objectsService.MfgDimensionsStability.dr["% Change Thickness Oven"] = DBNull.Value; gChangeOvenThickness  = string.Empty; }
                    CheckLimits("gChangeOvenThickness");

                    break;
                case "Freezer-L":
                    if (!(_objectsService.MfgDimensionsStability.dr["%change L - C1"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change L - C1"]; }
                    if (!(_objectsService.MfgDimensionsStability.dr["%change L - C2"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change L - C2"]; }
                    if (nCount == 2) { dtmp = dSum / (double)nCount; _objectsService.MfgDimensionsStability.dr["% Change Length Freezer"] = dtmp; gChangeFreezerLength  = dtmp.ToString("0.000"); } else { _objectsService.MfgDimensionsStability.dr["% Change Length Freezer"] = DBNull.Value; gChangeFreezerLength  = string.Empty; }
                    CheckLimits("gChangeFreezerLength");

                    break;
                case "Freezer-W":
                    if (!(_objectsService.MfgDimensionsStability.dr["%change W - C1"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change W - C1"]; }
                    if (!(_objectsService.MfgDimensionsStability.dr["%change W - C2"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change W - C2"]; }
                    if (nCount == 2) { dtmp = dSum / (double)nCount; _objectsService.MfgDimensionsStability.dr["% Change Width Freezer"] = dtmp; gChangeFreezerWidth  = dtmp.ToString("0.000"); } else { _objectsService.MfgDimensionsStability.dr["% Change Width Freezer"] = DBNull.Value; gChangeFreezerWidth  = string.Empty; }
                    CheckLimits("gChangeFreezerWidth");

                    break;
                case "Freezer-T":
                    if (!(_objectsService.MfgDimensionsStability.dr["%change T - C1"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change T - C1"]; }
                    if (!(_objectsService.MfgDimensionsStability.dr["%change T - C2"] == DBNull.Value)) { nCount += 1; dSum += (double)_objectsService.MfgDimensionsStability.dr["%change T - C2"]; }
                    if (nCount == 2) { dtmp = dSum / (double)nCount; _objectsService.MfgDimensionsStability.dr["% Change Thickness Freezer"] = dtmp; gChangeFreezerThickness  = dtmp.ToString("0.000"); } else { _objectsService.MfgDimensionsStability.dr["% Change Thickness Freezer"] = DBNull.Value; gChangeFreezerThickness  = string.Empty; }
                    CheckLimits("gChangeFreezerThickness");

                    break;
            }

        }

        public IActionResult OnPostGOven2_LF(string Name, string value)
        {

            string sName = Name;
            bool bl = false, bw = false, bt = false;
            int nCount = 0;
            double dSum = 0, dSumInit = 0, dSumFinal = 0;

            _objectsService.MfgDimensionsStability.bDataSetChanged = true;
            switch (sName)
            {
                case "gInit_H2_L1":  _objectsService.MfgDimensionsStability.dr["Initial H2 - L1"]=value; bl = true; break;
                case "gInit_H2_L2":  _objectsService.MfgDimensionsStability.dr["Initial H2 - L2"]=value; bl = true; break;
                case "gInit_H2_L3":  _objectsService.MfgDimensionsStability.dr["Initial H2 - L3"]=value; bl = true; break;
                case "gFinal_H2_L1": _objectsService.MfgDimensionsStability.dr[ "Final H2 - L1"]=value; bl = true; break;
                case "gFinal_H2_L2": _objectsService.MfgDimensionsStability.dr[ "Final H2 - L2"]=value; bl = true; break;
                case "gFinal_H2_L3": _objectsService.MfgDimensionsStability.dr[ "Final H2 - L3"]=value; bl = true; break;

                case "gInit_H2_W1": _objectsService.MfgDimensionsStability.dr["Initial H2 - W1"]=value; bw = true; break;
                case "gInit_H2_W2": _objectsService.MfgDimensionsStability.dr["Initial H2 - W2"]=value; bw = true; break;
                case "gInit_H2_W3": _objectsService.MfgDimensionsStability.dr["Initial H2 - W3"]=value; bw = true; break;
                case "gFinal_H2_W1":_objectsService.MfgDimensionsStability.dr[ "Final H2 - W1"]=value; bw = true; break;
                case "gFinal_H2_W2":_objectsService.MfgDimensionsStability.dr[ "Final H2 - W2"]=value; bw = true; break;
                case "gFinal_H2_W3":_objectsService.MfgDimensionsStability.dr[ "Final H2 - W3"]=value; bw = true; break;

                case "gInit_H2_T1":  _objectsService.MfgDimensionsStability.dr["Initial H2 - T1"]=value; bt = true; break;
                case "gInit_H2_T2":  _objectsService.MfgDimensionsStability.dr["Initial H2 - T2"]=value; bt = true; break;
                case "gInit_H2_T3":  _objectsService.MfgDimensionsStability.dr["Initial H2 - T3"]=value; bt = true; break;
                case "gInit_H2_T4":  _objectsService.MfgDimensionsStability.dr["Initial H2 - T4"]=value; bt = true; break;
                case "gInit_H2_T5":  _objectsService.MfgDimensionsStability.dr["Initial H2 - T5"]=value; bt = true; break;
                case "gFinal_H2_T1": _objectsService.MfgDimensionsStability.dr["Final H2 - T1"] = value; bt = true; break;
                case "gFinal_H2_T2": _objectsService.MfgDimensionsStability.dr[ "Final H2 - T2"]=value; bt = true; break;
                case "gFinal_H2_T3": _objectsService.MfgDimensionsStability.dr[ "Final H2 - T3"]=value; bt = true; break;
                case "gFinal_H2_T4": _objectsService.MfgDimensionsStability.dr[ "Final H2 - T4"]=value; bt = true; break;
                case "gFinal_H2_T5": _objectsService.MfgDimensionsStability.dr[ "Final H2 - T5"]=value; bt = true; break;
            }


            if (bl)
            {
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - L1"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - L1"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - L1"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - L1"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - L2"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - L2"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - L2"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - L2"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - L3"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - L3"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - L3"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - L3"]; }
                if (nCount == 3 && dSumInit > 0) {  _objectsService.MfgDimensionsStability.dr["%change L - H2"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else  _objectsService.MfgDimensionsStability.dr["%change L - H2"] = DBNull.Value;
                DimensionChanges("Oven-L");
            }
            else if (bw)
            {
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - W1"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - W1"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - W1"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - W1"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - W2"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - W2"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - W2"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - W2"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - W3"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - W3"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - W3"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - W3"]; }
                if (nCount == 3 && dSumInit > 0) {  _objectsService.MfgDimensionsStability.dr["%change W - H2"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else  _objectsService.MfgDimensionsStability.dr["%change W - H2"] = DBNull.Value;
                DimensionChanges("Oven-W");
            }
            else if (bt)
            {
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - T1"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - T1"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - T1"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - T1"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - T2"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - T2"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - T2"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - T2"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - T3"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - T3"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - T3"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - T3"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - T4"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - T4"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - T4"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - T4"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial H2 - T5"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final H2 - T5"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial H2 - T5"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final H2 - T5"]; }


                if (nCount == 5 && dSumInit > 0) {  _objectsService.MfgDimensionsStability.dr["%change T - H2"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else  _objectsService.MfgDimensionsStability.dr["%change T - H2"] = DBNull.Value;
                DimensionChanges("Oven-T");
            }

            _objectsService.MfgDimensionsStability.UpdateDataSet();
            return new JsonResult(new { message = Name, value });
        }

        public IActionResult OnPostGFreezer1_LF(string Name, string value)
        {
            string sName = Name;
            bool bl = false, bw = false, bt = false;
            int nCount = 0;
            double dSum = 0, dSumInit = 0, dSumFinal = 0;

            _objectsService.MfgDimensionsStability.bDataSetChanged = true;
            switch (sName)
            {
                case "gInit_C1_L1": _objectsService.MfgDimensionsStability.dr["Initial C1 - L1"]=value; bl = true; break;
                case "gInit_C1_L2": _objectsService.MfgDimensionsStability.dr["Initial C1 - L2"]=value; bl = true; break;
                case "gInit_C1_L3": _objectsService.MfgDimensionsStability.dr["Initial C1 - L3"]=value; bl = true; break;
                case "gFinal_C1_L1": _objectsService.MfgDimensionsStability.dr["Final C1 - L1"]=value; bl = true; break;
                case "gFinal_C1_L2": _objectsService.MfgDimensionsStability.dr["Final C1 - L2"]=value; bl = true; break;
                case "gFinal_C1_L3": _objectsService.MfgDimensionsStability.dr["Final C1 - L3"]=value; bl = true; break;

                case "gInit_C1_W1": _objectsService.MfgDimensionsStability.dr["Initial C1 - W1"]=value; bw = true; break;
                case "gInit_C1_W2": _objectsService.MfgDimensionsStability.dr["Initial C1 - W2"]=value; bw = true; break;
                case "gInit_C1_W3": _objectsService.MfgDimensionsStability.dr["Initial C1 - W3"]=value; bw = true; break;
                case "gFinal_C1_W1":  _objectsService.MfgDimensionsStability.dr["Final C1 - W1"]=value; bw = true; break;
                case "gFinal_C1_W2":  _objectsService.MfgDimensionsStability.dr["Final C1 - W2"]=value; bw = true; break;
                case "gFinal_C1_W3":  _objectsService.MfgDimensionsStability.dr["Final C1 - W3"]=value; bw = true; break;

                case "gInit_C1_T1": _objectsService.MfgDimensionsStability.dr["Initial C1 - T1"]=value; bt = true; break;
                case "gInit_C1_T2": _objectsService.MfgDimensionsStability.dr["Initial C1 - T2"]=value; bt = true; break;
                case "gInit_C1_T3": _objectsService.MfgDimensionsStability.dr["Initial C1 - T3"]=value; bt = true; break;
                case "gInit_C1_T4": _objectsService.MfgDimensionsStability.dr["Initial C1 - T4"]=value; bt = true; break;
                case "gInit_C1_T5": _objectsService.MfgDimensionsStability.dr["Initial C1 - T5"]=value; bt = true; break;
                case "gFinal_C1_T1": _objectsService.MfgDimensionsStability.dr["Final C1 - T1"]=value; bt = true; break;
                case "gFinal_C1_T2": _objectsService.MfgDimensionsStability.dr["Final C1 - T2"]=value; bt = true; break;
                case "gFinal_C1_T3": _objectsService.MfgDimensionsStability.dr["Final C1 - T3"]=value; bt = true; break;
                case "gFinal_C1_T4": _objectsService.MfgDimensionsStability.dr["Final C1 - T4"]=value; bt = true; break;
                case "gFinal_C1_T5": _objectsService.MfgDimensionsStability.dr["Final C1 - T5"]=value; bt = true; break;
            }


            if (bl)
            {
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - L1"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - L1"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - L1"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - L1"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - L2"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - L2"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - L2"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - L2"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - L3"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - L3"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - L3"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - L3"]; }
                if (nCount == 3 && dSumInit > 0) {  _objectsService.MfgDimensionsStability.dr["%change L - C1"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else  _objectsService.MfgDimensionsStability.dr["%change L - C1"] = DBNull.Value;
                DimensionChanges("Freezer-L");
            }
            else if (bw)
            {
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - W1"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - W1"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - W1"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - W1"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - W2"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - W2"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - W2"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - W2"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - W3"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - W3"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - W3"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - W3"]; }
                if (nCount == 3 && dSumInit > 0) {  _objectsService.MfgDimensionsStability.dr["%change W - C1"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else  _objectsService.MfgDimensionsStability.dr["%change W - C1"] = DBNull.Value;
                DimensionChanges("Freezer-W");
            }
            else if (bt)
            {
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - T1"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - T1"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - T1"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - T1"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - T2"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - T2"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - T2"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - T2"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - T3"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - T3"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - T3"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - T3"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - T4"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - T4"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - T4"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - T4"]; }
                if (!( _objectsService.MfgDimensionsStability.dr["Initial C1 - T5"] == DBNull.Value) && !( _objectsService.MfgDimensionsStability.dr["Final C1 - T5"] == DBNull.Value)) { nCount++; dSumInit += (double) _objectsService.MfgDimensionsStability.dr["Initial C1 - T5"]; dSumFinal += (double) _objectsService.MfgDimensionsStability.dr["Final C1 - T5"]; }


                if (nCount == 5 && dSumInit > 0) {  _objectsService.MfgDimensionsStability.dr["%change T - C1"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else  _objectsService.MfgDimensionsStability.dr["%change T - C1"] = DBNull.Value;
                DimensionChanges("Freezer-T");
            }

            _objectsService.MfgDimensionsStability.UpdateDataSet();
            return new JsonResult(new { message = Name, value });

        }

        public IActionResult OnPostGFreezer2_LF(string Name, string value)
        {

            string sName = Name;
            bool bl = false, bw = false, bt = false;
            int nCount = 0;
            double dSum = 0, dSumInit = 0, dSumFinal = 0;

            _objectsService.MfgDimensionsStability.bDataSetChanged = true;
            switch (sName)
            {
                case "gInit_C2_L1": _objectsService.MfgDimensionsStability.dr["Initial C2 - L1"]=value; bl = true; break;
                case "gInit_C2_L2": _objectsService.MfgDimensionsStability.dr["Initial C2 - L2"]=value; bl = true; break;
                case "gInit_C2_L3": _objectsService.MfgDimensionsStability.dr["Initial C2 - L3"]=value; bl = true; break;
                case "gFinal_C2_L1": _objectsService.MfgDimensionsStability.dr["Final C2 - L1"]=value; bl = true; break;
                case "gFinal_C2_L2": _objectsService.MfgDimensionsStability.dr["Final C2 - L2"]=value; bl = true; break;
                case "gFinal_C2_L3": _objectsService.MfgDimensionsStability.dr["Final C2 - L3"]=value; bl = true; break;

                case "gInit_C2_W1": _objectsService.MfgDimensionsStability.dr["Initial C2 - W1"]=value; bw = true; break;
                case "gInit_C2_W2": _objectsService.MfgDimensionsStability.dr["Initial C2 - W2"]=value; bw = true; break;
                case "gInit_C2_W3": _objectsService.MfgDimensionsStability.dr["Initial C2 - W3"]=value; bw = true; break;
                case "gFinal_C2_W1": _objectsService.MfgDimensionsStability.dr["Final C2 - W1"]=value; bw = true; break;
                case "gFinal_C2_W2": _objectsService.MfgDimensionsStability.dr["Final C2 - W2"]=value; bw = true; break;
                case "gFinal_C2_W3": _objectsService.MfgDimensionsStability.dr["Final C2 - W3"]=value; bw = true; break;

                case "gInit_C2_T1": _objectsService.MfgDimensionsStability.dr["Initial C2 - T1"]=value; bt = true; break;
                case "gInit_C2_T2": _objectsService.MfgDimensionsStability.dr["Initial C2 - T2"]=value; bt = true; break;
                case "gInit_C2_T3": _objectsService.MfgDimensionsStability.dr["Initial C2 - T3"]=value; bt = true; break;
                case "gInit_C2_T4": _objectsService.MfgDimensionsStability.dr["Initial C2 - T4"]=value; bt = true; break;
                case "gInit_C2_T5": _objectsService.MfgDimensionsStability.dr["Initial C2 - T5"]=value; bt = true; break;
                case "gFinal_C2_T1": _objectsService.MfgDimensionsStability.dr["Final C2 - T1"]=value; bt = true; break;
                case "gFinal_C2_T2": _objectsService.MfgDimensionsStability.dr["Final C2 - T2"]=value; bt = true; break;
                case "gFinal_C2_T3": _objectsService.MfgDimensionsStability.dr["Final C2 - T3"]=value; bt = true; break;
                case "gFinal_C2_T4": _objectsService.MfgDimensionsStability.dr["Final C2 - T4"]=value; bt = true; break;
                case "gFinal_C2_T5": _objectsService.MfgDimensionsStability.dr["Final C2 - T5"]=value; bt = true; break;
            }


            if (bl)
            {
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - L1"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - L1"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - L1"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - L1"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - L2"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - L2"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - L2"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - L2"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - L3"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - L3"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - L3"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - L3"]; }
                if (nCount == 3 && dSumInit > 0) { _objectsService.MfgDimensionsStability.dr["%change L - C2"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else _objectsService.MfgDimensionsStability.dr["%change L - C2"] = DBNull.Value;
                DimensionChanges("Freezer-L");
            }
            else if (bw)
            {
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - W1"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - W1"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - W1"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - W1"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - W2"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - W2"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - W2"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - W2"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - W3"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - W3"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - W3"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - W3"]; }
                if (nCount == 3 && dSumInit > 0) { _objectsService.MfgDimensionsStability.dr["%change W - C2"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else _objectsService.MfgDimensionsStability.dr["%change W - C2"] = DBNull.Value;
                DimensionChanges("Freezer-W");
            }
            else if (bt)
            {
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - T1"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - T1"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - T1"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - T1"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - T2"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - T2"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - T2"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - T2"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - T3"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - T3"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - T3"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - T3"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - T4"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - T4"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - T4"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - T4"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["Initial C2 - T5"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["Final C2 - T5"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["Initial C2 - T5"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["Final C2 - T5"]; }


                if (nCount == 5 && dSumInit > 0) { _objectsService.MfgDimensionsStability.dr["%change T - C2"] = 100.0 * (dSumFinal - dSumInit) / dSumInit; } else _objectsService.MfgDimensionsStability.dr["%change T - C2"] = DBNull.Value;
                DimensionChanges("Freezer-T");
            }

            _objectsService.MfgDimensionsStability.UpdateDataSet();
            return new JsonResult(new { message = Name, value });

        }

        public IActionResult OnPostGDateTime_LF(string value)
        {
            if (_objectsService.MfgDimensionsStability.dr["DateSampleIn"]==null)
                gTestDateTime = null;
             else
                _objectsService.MfgDimensionsStability.dr["DateSampleIn"]=value;

            _objectsService.MfgDimensionsStability.UpdateDataSet();
            return new JsonResult(new { message = value });
        }

        public IActionResult OnPostGWFDim_LF(string Name, string value)
        {

            string sName = Name;
            bool bl = false, bw = false, bt = false;
            int nCount = 0;
            double dSumInit = 0, dMax = double.MinValue, dSumFinal = 0;
            string sField; double dtmp;

            _objectsService.MfgDimensionsStability.bDataSetChanged = true;
            switch (sName)
            {
                case "gInit_WF_L1":
                    _objectsService.MfgDimensionsStability.dr["InitialWF-L1"]=value; bl = true; break;
                case "gInit_WF_L2":
                    _objectsService.MfgDimensionsStability.dr[ "InitialWF-L2"] = value; bl = true; break;
                case "gInit_WF_L3":
                    _objectsService.MfgDimensionsStability.dr["InitialWF-L3"] = value; bl = true; break;
                case "gFinal_WF_L1":
                    _objectsService.MfgDimensionsStability.dr["FinalWF-L1"] = value; bl = true; break;
                case "gFinal_WF_L2":
                    _objectsService.MfgDimensionsStability.dr[ "FinalWF-L2"] = value; bl = true; break;
                case "gFinal_WF_L3":
                    _objectsService.MfgDimensionsStability.dr["FinalWF-L3"] = value; bl = true; break;

                case "gInit_WF_W1":
                    _objectsService.MfgDimensionsStability.dr[ "InitialWF-W1"] = value; bw = true; break;
                case "gInit_WF_W2":
                    _objectsService.MfgDimensionsStability.dr["InitialWF-W2"] = value; bw = true; break;
                case "gInit_WF_W3":
                    _objectsService.MfgDimensionsStability.dr[ "InitialWF-W3"] = value; bw = true; break;
                case "gFinal_WF_W1":
                    _objectsService.MfgDimensionsStability.dr[ "FinalWF-W1"] = value; bw = true; break;
                case "gFinal_WF_W2":
                    _objectsService.MfgDimensionsStability.dr[ "FinalWF-W2"] = value; bw = true; break;
                case "gFinal_WF_W3":
                    _objectsService.MfgDimensionsStability.dr["FinalWF-W3"] = value; bw = true; break;
            }

            if (bl)
            {
                if (!(_objectsService.MfgDimensionsStability.dr["InitialWF-L1"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["FinalWF-L1"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["InitialWF-L1"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["FinalWF-L1"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["InitialWF-L2"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["FinalWF-L2"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["InitialWF-L2"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["FinalWF-L2"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["InitialWF-L3"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["FinalWF-L3"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["InitialWF-L3"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["FinalWF-L3"]; }

                if (nCount > 1)
                {
                    dtmp = 100.0 * (dSumFinal - dSumInit) / dSumInit; _objectsService.MfgDimensionsStability.dr["%ChangeLengthWF"] = dtmp; gChangeWFLength = dtmp.ToString(MfgDimStability.sOr);
                }
                else { _objectsService.MfgDimensionsStability.dr["%ChangeLengthWF"] = DBNull.Value; gChangeWFLength = string.Empty; }
            }
            else if (bw)
            {
                if (!(_objectsService.MfgDimensionsStability.dr["InitialWF-W1"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["FinalWF-W1"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["InitialWF-W1"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["FinalWF-W1"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["InitialWF-W2"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["FinalWF-W2"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["InitialWF-W2"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["FinalWF-W2"]; }
                if (!(_objectsService.MfgDimensionsStability.dr["InitialWF-W3"] == DBNull.Value) && !(_objectsService.MfgDimensionsStability.dr["FinalWF-W3"] == DBNull.Value)) { nCount++; dSumInit += (double)_objectsService.MfgDimensionsStability.dr["InitialWF-W3"]; dSumFinal += (double)_objectsService.MfgDimensionsStability.dr["FinalWF-W3"]; }

                if (nCount > 1)
                {
                    dtmp = 100.0 * (dSumFinal - dSumInit) / dSumInit; _objectsService.MfgDimensionsStability.dr["%ChangeWidthWF"] = dtmp; gChangeWFWidth = dtmp.ToString(MfgDimStability.sOr);
                }
                else
                {
                    _objectsService.MfgDimensionsStability.dr["%ChangeWidthWF"] = DBNull.Value; gChangeWFWidth = string.Empty;
                }

            }
            _objectsService.MfgDimensionsStability.UpdateDataSet();
            return new JsonResult(new { message = value }); 


        }

        public IActionResult OnPostGWFDepth_LF(string Name, string value)
        {

            string sName = Name;
            bool b1 = false, b2 = false, bt = false;
            int nCount1 = 0, nCount2 = 0;
            double dSum = 0, dMax = double.MinValue, dSumFinal = 0;
            string sField; double dtmp;

            _objectsService.MfgDimensionsStability.bDataSetChanged = true;
            switch (sName)
            {
                case "gSide1_Depth1": _objectsService.MfgDimensionsStability.dr["Loc1Depth1"]=value; b1 = true; break;
                case "gSide1_Depth2": _objectsService.MfgDimensionsStability.dr["Loc1Depth2"]=value; b1 = true; break;
                case "gSide1_Depth3": _objectsService.MfgDimensionsStability.dr["Loc1Depth3"]=value; b1 = true; break;
                case "gSide1_Depth4": _objectsService.MfgDimensionsStability.dr["Loc1Depth4"]=value; b1 = true; break;
                case "gSide1_Depth5": _objectsService.MfgDimensionsStability.dr["Loc1Depth5"]=value; b1 = true; break;

                case "gSide2_Depth1": _objectsService.MfgDimensionsStability.dr["Loc2Depth1"]=value; b2 = true; break;
                case "gSide2_Depth2": _objectsService.MfgDimensionsStability.dr["Loc2Depth2"]=value; b2 = true; break;
                case "gSide2_Depth3": _objectsService.MfgDimensionsStability.dr["Loc2Depth3"]=value; b2 = true; break;
                case "gSide2_Depth4": _objectsService.MfgDimensionsStability.dr["Loc2Depth4"]=value; b2 = true; break;
                case "gSide2_Depth5": _objectsService.MfgDimensionsStability.dr["Loc2Depth5"]=value; b2 = true; break;
            }

            for (int i = 1; i <= 5; i++)
            {
                sField = "Loc1Depth" + i.ToString();
                if (!(_objectsService.MfgDimensionsStability.dr[sField] == DBNull.Value))
                { nCount1++; dSum += (double)_objectsService.MfgDimensionsStability.dr[sField]; if (dMax < (double)_objectsService.MfgDimensionsStability.dr[sField]) dMax = (double)_objectsService.MfgDimensionsStability.dr[sField]; }

                sField = "Loc2Depth" + i.ToString();
                if (!(_objectsService.MfgDimensionsStability.dr[sField] == DBNull.Value)) { nCount2++; dSum += (double)_objectsService.MfgDimensionsStability.dr[sField]; if (dMax < (double)_objectsService.MfgDimensionsStability.dr[sField]) dMax = (double)_objectsService.MfgDimensionsStability.dr[sField]; }
            }

            if (nCount1 > 3 && nCount2 > 3)
            {
                dtmp = dSum / (double)(nCount1 + nCount2); _objectsService.MfgDimensionsStability.dr["AvgDepth"] = dtmp; gAvgDepth = dtmp.ToString(MfgDimStability.sOr);
                _objectsService.MfgDimensionsStability.dr["MaxDepth"] = dMax; gMaxDepth = dMax.ToString(MfgDimStability.sDef);
            }
            else
            {
                _objectsService.MfgDimensionsStability.dr["MaxDepth"] = _objectsService.MfgDimensionsStability.dr["AvgDepth"] = DBNull.Value; gMaxDepth = gAvgDepth = string.Empty;
            }

            _objectsService.MfgDimensionsStability.UpdateDataSet();
            return new JsonResult(new { message = value });
        }


        public IActionResult OnPostGWF_Info_LF(string Name, string value)
        {
            string sName = Name;
            switch (sName)
            {
                case "gDeviation":
                    _objectsService.MfgDimensionsStability.dr["DevFromTable"]=value; break;
                case "gDevType":
                    if (_objectsService.MfgDimensionsStability.dr["DevType"] == null) _objectsService.MfgDimensionsStability.dr["DevType"] = DBNull.Value; else 
                        _objectsService.MfgDimensionsStability.dr["DevType"] = value; break;
            }

            _objectsService.MfgDimensionsStability.UpdateDataSet();
            return new JsonResult(new { message = Name,value });
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

            (_objectsService.MfgInProcess, _objectsService.MfgFinishedGoods,_objectsService.MfgDimensionsStability,_objectsService.MfgPlantsData,_objectsService.MfgJetMixing) = _objectsService.MfgHome.GetAllMfgData(_objectsService.MfgInProcess, _objectsService.MfgFinishedGoods,_objectsService.MfgDimensionsStability, _objectsService.MfgPlantsData, _objectsService.MfgJetMixing);
            SetView();

            // Enable/disable buttons based on the index
            gDataSetNextIsEnabled = _objectsService.Cbfile.iIDMfgIndex < (_objectsService.MfgHome.dt.Rows.Count - 1);
            gDataSetPrevIsEnabled = _objectsService.Cbfile.iIDMfgIndex > 0;
        }

        //public IActionResult gDateTime_GF(object sender, EventArgs e)
        //{

        //    System.Windows.Forms.DateTimePicker dtp = sender as System.Windows.Forms.DateTimePicker;
        //    if (dtp.CustomFormat == sNull) { dtp.Value = DateTime.Now; dtp.CustomFormat = sDateTimeFormat; }

        //}

        /*
        if (bl)
        {
            if (!(dr["Initial H1 - L1"] == DBNull.Value) && !(dr["Final H1 - L1"] == DBNull.Value) && (double)dr["Initial H1 - L1"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - L1"] / (double)dr["Initial H1 - L1"] - 1.0; }
            if (!(dr["Initial H1 - L2"] == DBNull.Value) && !(dr["Final H1 - L2"] == DBNull.Value) && (double)dr["Initial H1 - L2"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - L2"] / (double)dr["Initial H1 - L2"] - 1.0; }
            if (!(dr["Initial H1 - L3"] == DBNull.Value) && !(dr["Final H1 - L3"] == DBNull.Value) && (double)dr["Initial H1 - L3"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - L3"] / (double)dr["Initial H1 - L3"] - 1.0; }

            if (nCount > 0) { dr["%change L - H1"] = dSum / (double)nCount; DimensionChanges("Oven-L"); } else dr["%change L - H1"] = DBNull.Value;
        }
        else if (bw)
        {
            if (!(dr["Initial H1 - W1"] == DBNull.Value) && !(dr["Final H1 - W1"] == DBNull.Value) && (double)dr["Initial H1 - W1"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - W1"] / (double)dr["Initial H1 - W1"] - 1.0; }
            if (!(dr["Initial H1 - W2"] == DBNull.Value) && !(dr["Final H1 - W2"] == DBNull.Value) && (double)dr["Initial H1 - W2"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - W2"] / (double)dr["Initial H1 - W2"] - 1.0; }
            if (!(dr["Initial H1 - W3"] == DBNull.Value) && !(dr["Final H1 - W3"] == DBNull.Value) && (double)dr["Initial H1 - W3"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - W3"] / (double)dr["Initial H1 - W3"] - 1.0; }

            if (nCount > 0) { dr["%change W - H1"] = dSum / (double)nCount; DimensionChanges("Oven-W"); } else dr["%change W - H1"] = DBNull.Value;
        }

        else if (bt)
        {
            if (!(dr["Initial H1 - T1"] == DBNull.Value) && !(dr["Final H1 - T1"] == DBNull.Value) && (double)dr["Initial H1 - T1"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - T1"] / (double)dr["Initial H1 - T1"] - 1.0; }
            if (!(dr["Initial H1 - T2"] == DBNull.Value) && !(dr["Final H1 - T2"] == DBNull.Value) && (double)dr["Initial H1 - T2"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - T2"] / (double)dr["Initial H1 - T2"] - 1.0; }
            if (!(dr["Initial H1 - T3"] == DBNull.Value) && !(dr["Final H1 - T3"] == DBNull.Value) && (double)dr["Initial H1 - T3"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - T3"] / (double)dr["Initial H1 - T3"] - 1.0; }
            if (!(dr["Initial H1 - T4"] == DBNull.Value) && !(dr["Final H1 - T4"] == DBNull.Value) && (double)dr["Initial H1 - T4"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - T4"] / (double)dr["Initial H1 - T4"] - 1.0; }
            if (!(dr["Initial H1 - T5"] == DBNull.Value) && !(dr["Final H1 - T5"] == DBNull.Value) && (double)dr["Initial H1 - T5"] > 0)
            { nCount += 1; dSum += (double)dr["Final H1 - T5"] / (double)dr["Initial H1 - T5"] - 1.0; }
            if (nCount > 0) { dr["%change T - H1"] = dSum / (double)nCount; DimensionChanges("Oven-T"); } else dr["%change T - H1"] = DBNull.Value;
        }
        */


    }
}