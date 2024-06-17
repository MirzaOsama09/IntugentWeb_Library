using IntugentClassLbrary.Classes;
using IntugentClassLibrary.Pages.Mfg;
using IntugentClassLibrary.Utilities;
using IntugentWebApp.Controllers.Mfg;
using IntugentWebApp.Controllers.RnD;
using IntugentWebApp.Models;
using IntugentWebApp.Utilities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data;
using System.Data.SqlClient;
using System.Reflection.Metadata;
using System.Text;

namespace IntugentWebApp.Pages.Mfg_Group
{
    [BindProperties]
    public class HomeModel : PageModel
    {
       
        public string gLoc1 { get; set; }
      
        public DataView gProd1 { get; set; }
     
        public DataView gTestStat1 { get; set; }
     
        public DataView gAgedRValue { get;  set; }
   
        public DataView gDimStability { get;  set; }
   
        public DataView gRunType1 { get;  set; }
 
        public string gProd1SelectedValue { get;  set; }
      
        public DateTime? gMfgDate1 { get;  set; }
    
        public DateTime? gMfgDate2 { get;  set; }
  
        public int gTestStat1SelectedValue { get;  set; }
     
        public int gAgedRValueSelectedValue { get;  set; }
       
        public int gDimStabilitySelectedValue { get;  set; }
       
        public int gRunType1SelectedValue { get;  set; }
      
        public DataView gMfgSearch { get;  set; }
        
        public int gMfgSearchSelectedIndex { get;  set; }

        private readonly ObjectsService _objectsService;

        public HomeModel(ObjectsService objectsService)
        {
            _objectsService = objectsService;
        }

        public async void OnGet()
        {
            try
            {
                SetData();
                Startup(_objectsService.MfgHome);
                // selectedDatasetID = HttpContext.Session.GetString("selectedDatasetID");
                /* searchResults = await MfgHomeController.GetSearchResults();

                 searchParams.ProductCode = _objectsService.CLists.dvComProdAll;
                 searchParams.Group = _objectsService.CDefualts.sLocation;
                 MyProperty = _objectsService.CDefualts.sProdMfg;
                 selectedDatasetID = _objectsService.CLists.drEmployee["MfgIDSelected"].ToString();*/

                //TempData["ErrorOnServer"] = gMfgDate1 + " -------- "; 
            }
            catch (Exception ex)
            {
                TempData["ErrorOnServer"] = ex.Message;
            }
        }

        private void SetData()
        {
            #region Mfg specific lists
            if (_objectsService.CDefualts.IDLocation != 3)
            {
                gLoc1 = _objectsService.CDefualts.sLocation;
                gProd1 = _objectsService.CLists.dvComProdAll;

                _objectsService.CLists.dvLists = _objectsService.CLists.dtLists.DefaultView;
                _objectsService.CLists.dvLists.RowFilter = "sList = 'Testing Status Mfg'";  //Testing Status
                _objectsService.CLists.dvTestingStat = _objectsService.CLists.dvLists.ToTable().DefaultView;
                gTestStat1 = _objectsService.CLists.dvTestingStat;

                _objectsService.CLists.dvLists.RowFilter = "sList = 'Aged R Value Mfg'";
                _objectsService.CLists.dvTestingStat = _objectsService.CLists.dvLists.ToTable().DefaultView;
                gAgedRValue = _objectsService.CLists.dvTestingStat;

                _objectsService.CLists.dvLists.RowFilter = "sList = 'Dim Stability Mfg'";
                _objectsService.CLists.dvTestingStat = _objectsService.CLists.dvLists.ToTable().DefaultView;
                gDimStability = _objectsService.CLists.dvTestingStat;

                _objectsService.CLists.dvLists.RowFilter = "sList = 'Run Type Mfg'";
                _objectsService.CLists.dvRunType = _objectsService.CLists.dvLists.ToTable().DefaultView;
                gRunType1 = _objectsService.CLists.dvRunType;



                _objectsService.CLists.dvLists.RowFilter = "sList = 'Run Type Mfg' AND ID <> 50";
                _objectsService.CLists.dvRunType2 = _objectsService.CLists.dvLists.ToTable().DefaultView;

                _objectsService.CLists.dvLists.RowFilter = "sList = 'Surfactant'";
                _objectsService.CLists.dvSurfactants = _objectsService.CLists.dvLists.ToTable().DefaultView;

                _objectsService.CLists.dvLists.RowFilter = "sList = 'Pour Table Layout'";
                _objectsService.CLists.dvLayout = _objectsService.CLists.dvLists.ToTable().DefaultView;

                _objectsService.CLists.dvLists.RowFilter = "sList = 'Paper'";
                _objectsService.CLists.dvPaper = _objectsService.CLists.dvLists.ToTable().DefaultView;
                _objectsService.CLists.dvPaper.Sort = "sName";

                _objectsService.CLists.dvLists.RowFilter = "sList = 'Shift'";
                _objectsService.CLists.dvShift = _objectsService.CLists.dvLists.ToTable().DefaultView;

                _objectsService.CLists.dvEmployees = _objectsService.CLists.dtLoc.DefaultView;
                if (_objectsService.CDefualts.IDLocation == 4) _objectsService.CLists.dvEmployees.RowFilter = "IDLocation <> 3 and IDLocation <> 4";
                else _objectsService.CLists.dvEmployees.RowFilter = "IDLocation = " + _objectsService.CDefualts.IDLocation.ToString();  //Employees
                _objectsService.CLists.dvEmployeesMfg = _objectsService.CLists.dvEmployees.ToTable().DefaultView;

                _objectsService.CLists.dvLists.RowFilter = "sList = 'ProcessCheckType'"; 
                _objectsService.CLists.dvPCType = _objectsService.CLists.dvLists.ToTable().DefaultView;

                _objectsService.CLists.dvLists.RowFilter = "sList = 'ProcessCheckTopBoard'";
                _objectsService.CLists.dvPCTopBoard = _objectsService.CLists.dvLists.ToTable().DefaultView;
                _objectsService.CLists.dvLists.RowFilter = "sList = 'ProcessCheckBottomBoard'";
                _objectsService.CLists.dvPCBottomBoard = _objectsService.CLists.dvLists.ToTable().DefaultView;
                _objectsService.CLists.dvLists.RowFilter = "sList = 'ProcessCheckPerferation'";
                _objectsService.CLists.dvPCPerferation = _objectsService.CLists.dvLists.ToTable().DefaultView;
                _objectsService.CLists.dvLists.RowFilter = "sList = 'ProcessCheckFlipper'";
                _objectsService.CLists.dvPCFlipper = _objectsService.CLists.dvLists.ToTable().DefaultView;

                _objectsService.CLists.dvLists.RowFilter = "sList = 'ProcessCheckFacerAdh '";
                _objectsService.CLists.dvPCFacerAdh = _objectsService.CLists.dvLists.ToTable().DefaultView;
                _objectsService.CLists.dvLists.RowFilter = "sList = 'ProcessCheckEdgeCut'";
                _objectsService.CLists.dvPCEdgeCut = _objectsService.CLists.dvLists.ToTable().DefaultView;
                _objectsService.CLists.dvLists.RowFilter = "sList = 'ProcessCheckHooderQuality'";
                _objectsService.CLists.dvPCHooderQuality = _objectsService.CLists.dvLists.ToTable().DefaultView;
                _objectsService.CLists.dvLists.RowFilter = "sList = 'ProcessCheckBoardQuality'";
                _objectsService.CLists.dvPCBoardQuality = _objectsService.CLists.dvLists.ToTable().DefaultView;
            }
            #endregion
        }

        public void Startup(MfgHome mfgHome)
        {
            if (_objectsService.CLists.drEmployee["Mfg Product Code"] == DBNull.Value) gProd1SelectedValue = _objectsService.CDefualts.sProdMfgAll; else gProd1SelectedValue = (string)_objectsService.CLists.drEmployee["Mfg Product Code"];
            if (_objectsService.CLists.drEmployee["MfgDate1"] == DBNull.Value) gMfgDate1 = null; else gMfgDate1 = (DateTime)_objectsService.CLists.drEmployee["MfgDate1"];
            if (_objectsService.CLists.drEmployee["MfgDate2"] == DBNull.Value) gMfgDate2 = null; else gMfgDate2 = (DateTime)_objectsService.CLists.drEmployee["MfgDate2"];
            if (_objectsService.CLists.drEmployee["MfgIDTestingStatus"] == DBNull.Value) gTestStat1SelectedValue = _objectsService.CDefualts.iMfgTestingStat; else gTestStat1SelectedValue = (int)_objectsService.CLists.drEmployee["MfgIDTestingStatus"];
            if (_objectsService.CLists.drEmployee["MfgIDAgedTesting"] == DBNull.Value) gAgedRValueSelectedValue = _objectsService.CDefualts.iMfgAgedRValue; else gAgedRValueSelectedValue = (int)_objectsService.CLists.drEmployee["MfgIDAgedTesting"];
            if (_objectsService.CLists.drEmployee["MfgIDDimStability"] == DBNull.Value) gDimStabilitySelectedValue = _objectsService.CDefualts.iMfgDimStability; else gDimStabilitySelectedValue = (int)_objectsService.CLists.drEmployee["MfgIDDimStability"];
            if (_objectsService.CLists.drEmployee["MfgIDRunType"] == DBNull.Value) gRunType1SelectedValue = _objectsService.CDefualts.iMfgRunType; else gRunType1SelectedValue = (int)_objectsService.CLists.drEmployee["MfgIDRunType"];
            
            if (mfgHome.SearchMfgDB() && mfgHome.dt.Rows.Count > 0)
            {
                _objectsService.Cbfile.iIDMfgIndex = 0; 
                _objectsService.Cbfile.iIDMfg = (int)mfgHome.dt.Rows[0]["ID4ALL"];

                if (_objectsService.CLists.drEmployee["MfgIDSelected"] != DBNull.Value)
                {
                    int idTemp = (int)_objectsService.CLists.drEmployee["MfgIDSelected"];
                    for (int i = 0; i < mfgHome.dt.Rows.Count; i++)
                    {
                        if ((int)mfgHome.dt.Rows[i]["ID4ALL"] == idTemp)
                        {
                            _objectsService.Cbfile.iIDMfgIndex = i;
                            _objectsService.Cbfile.iIDMfg = (int)mfgHome.dt.Rows[i]["ID4ALL"];
                        }
                    }
                }
                gMfgSearchSelectedIndex = _objectsService.Cbfile.iIDMfgIndex;
                gMfgSearch = mfgHome.dt.DefaultView;
            }
            else
            {
                gMfgSearch = mfgHome.dt.DefaultView;
            }
        }

        
        public IActionResult OnPostDatasetSelected(int gMfgSearchSelectedIndex, int gMfgSearchRowsCount, int gMfgSelectedDatasetID)
        {
            /*_objectsService.CLists.drEmployee["MfgIDSelected"] = id;
            CLists_UpdateEmployee.UpdateEmployee(_objectsService.CLists);*/

            if (!_objectsService.Cbfile.bCanSwitchRecord)
            {
                gMfgSearchSelectedIndex = _objectsService.Cbfile.iIDMfgIndex;
                //MessageBox.Show(Cbfile.sNoRecSwitchMsg, Cbfile.sAppName);
                return new JsonResult(new { message = "canswitch: " });
            }

            if (gMfgSearchSelectedIndex < 0 || gMfgSearchSelectedIndex > gMfgSearchRowsCount - 1) return new JsonResult(new { message = "idx < 0: " + gMfgSearchSelectedIndex + " -- " + gMfgSearchRowsCount });
            if (_objectsService.MfgHome.dt.Rows[gMfgSearchSelectedIndex]["ID4ALL"] == DBNull.Value)
            {
                // sMsg = "The selected dataset does not have a valid ID and can not be selected.";
                // MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Stop);
                // EnableMfgPages(false);
                return Page();
            }


            ///////////////// this is index of gMFGSearch dataview not dataset ID
            _objectsService.Cbfile.iIDMfgIndex = gMfgSearchSelectedIndex;

            _objectsService.Cbfile.iIDMfg = gMfgSelectedDatasetID;
            //EnableMfgPages(true);
            (_objectsService.MfgInProcess,_objectsService.MfgFinishedGoods, _objectsService.MfgDimensionsStability) = _objectsService.MfgHome.GetAllMfgData(_objectsService.MfgInProcess,_objectsService.MfgFinishedGoods, _objectsService.MfgDimensionsStability);

            return new JsonResult(new { message = "Dataset selected: " + gMfgSearchSelectedIndex + " -- " + gMfgSelectedDatasetID });
        }

        public IActionResult OnPostSearchDB_Click()
        {            
            if (gProd1SelectedValue == null) _objectsService.CLists.drEmployee["Mfg Product Code"] = DBNull.Value; else _objectsService.CLists.drEmployee["Mfg Product Code"] = gProd1SelectedValue;
            if (gMfgDate1 == null) _objectsService.CLists.drEmployee["MfgDate1"] = DBNull.Value; else _objectsService.CLists.drEmployee["MfgDate1"] = gMfgDate1;
            if (gMfgDate2 == null) _objectsService.CLists.drEmployee["MfgDate2"] = DBNull.Value; else _objectsService.CLists.drEmployee["MfgDate2"] = gMfgDate2;
            if (gTestStat1SelectedValue == null) _objectsService.CLists.drEmployee["MfgIDTestingStatus"] = DBNull.Value; else _objectsService.CLists.drEmployee["MfgIDTestingStatus"] = gTestStat1SelectedValue;

            if (gAgedRValueSelectedValue == null) _objectsService.CLists.drEmployee["MfgIDAgedTesting"] = DBNull.Value; else _objectsService.CLists.drEmployee["MfgIDAgedTesting"] = gAgedRValueSelectedValue;
            if (gDimStabilitySelectedValue == null) _objectsService.CLists.drEmployee["MfgIDDimStability"] = DBNull.Value; else _objectsService.CLists.drEmployee["MfgIDDimStability"] = gDimStabilitySelectedValue;
            if (gRunType1SelectedValue == null)
                _objectsService.CLists.drEmployee["MfgIDRunType"] = DBNull.Value;
            else
                _objectsService.CLists.drEmployee["MfgIDRunType"] = gRunType1SelectedValue;

            if (_objectsService.MfgHome.SearchMfgDB() && _objectsService.MfgHome.dt.Rows.Count > 0)
            {

                _objectsService.Cbfile.iIDMfgIndex = 0; _objectsService.Cbfile.iIDMfg = (int)_objectsService.MfgHome.dt.Rows[0]["ID4ALL"];
                gMfgSearch = _objectsService.MfgHome.dt.DefaultView;
                gMfgSearchSelectedIndex = _objectsService.Cbfile.iIDMfgIndex;

                CLists_UpdateEmployee.UpdateEmployee(_objectsService.CLists);
            }
            else
            {
                gMfgSearch = _objectsService.MfgHome.dt.DefaultView;


            }
            return RedirectToPage();
        }
        /*
                public async void OnPostParams()
                {
                    try
                    {
                        //selectedDatasetID = HttpContext.Session.GetString("selectedDatasetID");
                        searchResults = await MfgHomeController.GetSearchResults(searchParams);
                    }
                    catch (Exception ex)
                    {
                        TempData["ErrorOnServer"] = ex.Message;
                    }
                }

                public async Task<IActionResult> OnPostCopyDataAsync()
                {
                    List<MFGSearchResult> items = await MfgHomeController.GetSearchResults(searchParams);
                    var sb = new StringBuilder();

                    // Convert data to CSV format

                    foreach (MFGSearchResult item in items)
                    {
                        sb.AppendLine($"{item.ID}, {item.GreenBoardTimeStamp}");
                    }

                    var data = sb.ToString();
                    TempData["ErrorOnServer"] = data + ".............";
                    return new JsonResult(new {message = items.Count });
                }*/
    }
}
