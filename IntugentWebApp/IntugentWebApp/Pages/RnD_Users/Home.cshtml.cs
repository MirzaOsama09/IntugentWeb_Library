using IntugentClassLbrary.Classes;
using IntugentWebApp.Controllers.RnD;
using IntugentWebApp.Models;
using IntugentWebApp.Utilities;
using Microsoft.AspNetCore.Mvc;
using System.Text;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;
using System.Data;
using IntugentClassLibrary.Utilities;

namespace IntugentWebApp.Pages.RnD_Users
{
    [BindProperties]
    public class HomeModel : PageModel
    {
        public string gLoc1 { get; set; }
        public DataView gTestStat1{get;set;}
        public int gTestStat1SelectedValue { get;set;}
        public DataView gProd1 { get; set; }
        public string gProd1SelectedValue { get; set; }
        public DataView gStudyType { get; set; }
        public int gStudyTypeSelectedValue { get; set; }
        
        public DateTime? gRndDate1 { get; set; }
        public DateTime? gRndDate2 { get; set; }
        public string? selectedDatasetID { get; set; }
        public string gRNDNameSearch { get; set; }
        public DataView gRNDSearch {  get; set; }
        public int gRNDSearchSelectedIndex {  get; set; }

        public List<RNDSearchResult>? searchResults { get; set; }

        private readonly ObjectsService _objectsService;
        public HomeModel(ObjectsService objectsService)
        {
            _objectsService = objectsService;
            Startup();
        }
        public async void OnGet()
        {
            //_objectsService.RNDHome.GetNewDataset();
            try
            {
                selectedDatasetID = _objectsService.CLists.drEmployee["MfgIDSelected"].ToString();
                searchResults = await RndHomeController.GetSearchResults();
            }
            catch (Exception ex)
            {
                TempData["ErrorOnServer"] = ex.Message;
            }

            gRNDSearchSelectedIndex = _objectsService.RNDHome.indSet;
            if (gRNDSearchSelectedIndex >= 0)
                //   gRNDSearch.ScrollIntoView(gRNDSearch.Items[gRNDSearchSelectedIndex]);

                if (_objectsService.CLists.drEmployee == null) return;


            //           SetView();
        }

        public IActionResult OnPostCheckboxSelected(string id)
        {

            _objectsService.CLists.drEmployee["MfgIDSelected"] = id;
            CLists_UpdateEmployee.UpdateEmployee(_objectsService.CLists);

            return new JsonResult(new { message = "Dataset selected: " + id });
        }



        public void SetView()
        {
            //            gRNDSearch.SelectedValue = RND.IdSet;

        }

        public IActionResult OnPostGSearchDataSets_Click()
        {

            if (gRndDate1 == null) _objectsService.CLists.drEmployee["RndDate1"] = DBNull.Value; else _objectsService.CLists.drEmployee["RndDate1"] = gRndDate1;
            if (gRndDate2 == null) _objectsService.CLists.drEmployee["RndDate2"] = DBNull.Value; else _objectsService.CLists.drEmployee["RndDate2"] = gRndDate2;
            if (gProd1SelectedValue == null) _objectsService.CLists.drEmployee["Rnd Product Code"] = DBNull.Value; else _objectsService.CLists.drEmployee["Rnd Product Code"] = gProd1SelectedValue;
            if (gTestStat1SelectedValue == null) _objectsService.CLists.drEmployee["RndIDTestingStatus"] = DBNull.Value; else _objectsService.CLists.drEmployee["RndIDTestingStatus"] = gTestStat1SelectedValue;
            if (gStudyTypeSelectedValue == null) _objectsService.CLists.drEmployee["RndIDStudyType"] = DBNull.Value; else _objectsService.CLists.drEmployee["RndIDStudyType"] = gStudyTypeSelectedValue;
            if (String.IsNullOrEmpty(gRNDNameSearch)) _objectsService.CLists.drEmployee["RNDNameSearch"] = DBNull.Value; else _objectsService.CLists.drEmployee["RNDNameSearch"] = gRNDNameSearch;


            if (SearchRNDDB() && _objectsService.RNDHome.dt.Rows.Count > 0)
            {
                _objectsService.RNDHome.indSet = 0; _objectsService.RNDHome.IdSet = (int)_objectsService.RNDHome.dt.Rows[0]["ID"];
                gRNDSearch = _objectsService.RNDHome.dt.DefaultView;
                gRNDSearchSelectedIndex = _objectsService.RNDHome.indSet;

                //                gRNDSearch.ScrollIntoView(gRNDSearch.Items[gRNDSearch.SelectedIndex]);

                _objectsService.CLists.UpdateEmployee();
            }
            else
            {
                gRNDSearch = _objectsService.RNDHome.dt.DefaultView; _objectsService.RNDHome.EnableRNDPages(false);
            }
            return new JsonResult(new {message=true});
        }

        public bool SearchRNDDB()
        {
            //           DataTable dt1 = new DataTable();

            string sql = GetSearchCriteria();

            if (!string.IsNullOrEmpty(sql)) sql = _objectsService.RNDHome.sqlSearchDS + " Where " + sql; else sql = _objectsService.RNDHome.sqlSearchDS;
            sql = sql + " Order by DateDSCreated DESC ";

            try
            {
                _objectsService.RNDHome.da = new SqlDataAdapter(sql, _objectsService.Cbfile.conAZ);
                _objectsService.RNDHome.da.SelectCommand.Parameters.AddWithValue("@sParam1", _objectsService.RNDHome.sParamValue1);
                //                dt1.Clear();
                //               int itmp = da.Fill(dt1);
                _objectsService.RNDHome.dt.Clear();
                int itmp = _objectsService.RNDHome.da.Fill(_objectsService.RNDHome.dt);
                if (itmp < 1)
                {
                   // sMsg = ("No Dataset was found to meet the search criteria.  Check the network connection and/or relax the search criteria");
                   // if (!bInit) MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Information);
                    return false;
                }
                //               dr = dt.Rows[0];
                //                dt = dt1.Copy();
            }
            catch (SqlException ex)
            {
               // sMsg = "Error in searching for the datasets for editing";
              //  MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Stop);
               // System.Diagnostics.Trace.TraceError(sMsg + "\n\n" + ex.Message);
               // CTelClient.TelException(ex, sMsg);

                return false;
            }
            return true;

        }


        public IActionResult OnPostGNewDataSetRND_Click()
        {
            int idOld = _objectsService.RNDHome.IdSet;

            if (_objectsService.RNDHome.GetNewDataset())   //add new dataset to the list of search box.
            {
                _objectsService.RNDHome.dr = _objectsService.RNDHome.dt.NewRow();
                _objectsService.RNDHome.dr["ID"] = _objectsService.RNDHome.IdSet;
                _objectsService.RNDHome.dr["DateDSCreated"] = _objectsService.RNDHome.drS["DateDSCreated"];
                _objectsService.RNDHome.dt.Rows.InsertAt(_objectsService.RNDHome.dr, 0);
                gRNDSearch = _objectsService.RNDHome.dt.DefaultView;
                gRNDSearchSelectedIndex = 0;
              //  gRNDSearch.ScrollIntoView(gRNDSearch.Items[gRNDSearchSelectedIndex]);
            }
            return new JsonResult(true);
        }

        public IActionResult OnPostGSearchLF(string Name)
        {
            string sName = Name;
            //            bDataSetChanged = true;
            switch (sName)
            {
                //                case "gRndProp1Value1": SetDoubleField(gRndProp1Value1.Text, "RndProp1Value1"); break;
                //                case "gRndProp1Value2": SetDoubleField(gRndProp1Value2.Text, "RndProp1Value2"); break;
                //                case "gRndProp2Value1": SetDoubleField(gRndProp2Value1.Text, "RndProp2Value1"); break;
                //               case "gRndProp2Value2": SetDoubleField(gRndProp2Value2.Text, "RndProp2Value2"); break;
                //               case "gRndProp1": if (gRndProp1.SelectedValue == null) _objectsService.CLists.drEmployee["RndProp1"] = DBNull.Value; else _objectsService.CLists.drEmployee["RndProp1"] = (string) gRndProp1.SelectedValue; break;
                //               case "gRndProp2": if (gRndProp1.SelectedValue == null) _objectsService.CLists.drEmployee["RndProp2"] = DBNull.Value; else _objectsService.CLists.drEmployee["RndProp2"] = (string) gRndProp2.SelectedValue; break;
                case "gRndDate1": if (gRndDate1 == null) _objectsService.CLists.drEmployee["RndDate1"] = DBNull.Value; else _objectsService.CLists.drEmployee["RndDate1"] = gRndDate1; break;
                case "gRndDate2": if (gRndDate2 == null) _objectsService.CLists.drEmployee["RndDate2"] = DBNull.Value; else _objectsService.CLists.drEmployee["RndDate2"] = gRndDate2; break;
                case "gProd1": if (gProd1SelectedValue == null) _objectsService.CLists.drEmployee["Rnd Product Code"] = DBNull.Value; else _objectsService.CLists.drEmployee["Rnd Product Code"] = gProd1SelectedValue; break;
                case "gTestStat1": if (gTestStat1SelectedValue == null) _objectsService.CLists.drEmployee["RndIDTestingStatus"] = DBNull.Value; else _objectsService.CLists.drEmployee["RndIDTestingStatus"] = gTestStat1SelectedValue; break;
                case "gStudyType": if (gStudyTypeSelectedValue == null) _objectsService.CLists.drEmployee["RndIDStudyType"] = DBNull.Value; else _objectsService.CLists.drEmployee["RndIDStudyType"] = gStudyTypeSelectedValue; break;
                case "gRNDNameSearch": if (String.IsNullOrEmpty(gRNDNameSearch)) _objectsService.CLists.drEmployee["RNDNameSearch"] = DBNull.Value; else _objectsService.CLists.drEmployee["RNDNameSearch"] = gRNDNameSearch; break;

            }
            _objectsService.CLists.UpdateEmployee();
          return new JsonResult(true);
        }

        #region Set Fields

        public void SetDoubleField(string sText, string sField)
        {
            double dtmp;

            if (String.IsNullOrEmpty(sText)) _objectsService.CLists.drEmployee[sField] = DBNull.Value;
            else
            {
                if (double.TryParse(sText, out dtmp)) _objectsService.CLists.drEmployee[sField] = dtmp;
                else _objectsService.CLists.drEmployee[sField] = DBNull.Value;
            }

            _objectsService.CLists.UpdateEmployee();
        }

        public void SetIntField(string sText, string sField)
        {
            int itmp;

            if (String.IsNullOrEmpty(sText)) _objectsService.RNDHome.dr[sField] = DBNull.Value;
            else
            {
                if (int.TryParse(sText, out itmp)) _objectsService.RNDHome.dr[sField] = itmp;
                else _objectsService.RNDHome.dr[sField] = DBNull.Value;
            }

        }

        public string SetDoubleTextField(string sField, string sForm = "" )
        {
            string s = String.Empty;
            if (!(_objectsService.CLists.drEmployee[sField] == DBNull.Value)) s = ((double)_objectsService.CLists.drEmployee[sField]).ToString(sForm);
            return s;
        }

        public string SetIntTextField(string sField)
        {
            string s = String.Empty;
            if (!(_objectsService.RNDHome.dr[sField] == DBNull.Value)) s = ((int)_objectsService.RNDHome.dr[sField]).ToString();
            return s;
        }

        public IActionResult OnPostGRNDDatasetsLF()
        {

            int iOldSet = _objectsService.RNDHome.IdSet;
            if (gRNDSearchSelectedIndex >= 0)
            {
                _objectsService.RNDHome.IdSet = (int)_objectsService.RNDHome.dt.Rows[gRNDSearchSelectedIndex]["ID"];
                if (_objectsService.RNDHome.GetDataSet())
                {
                    _objectsService.RNDHome.indSet = gRNDSearchSelectedIndex; SetView(); _objectsService.CLists.drEmployee["RndIDSelected"] = _objectsService.RNDHome.IdSet; _objectsService.CLists.UpdateEmployee();
                }
                else _objectsService.RNDHome.IdSet = iOldSet;
            }
           return new JsonResult(true);

        }

        public IActionResult OnPostGRNDDatasetsSC()
        {
            int iOldSet = _objectsService.RNDHome.IdSet;
            if (gRNDSearchSelectedIndex >= 0)
            {
                _objectsService.RNDHome.IdSet = (int)_objectsService.RNDHome.dt.Rows[gRNDSearchSelectedIndex]["ID"];
                if (_objectsService.RNDHome.GetDataSet())
                {
                    _objectsService.RNDHome.indSet = gRNDSearchSelectedIndex; SetView(); _objectsService.CLists.drEmployee["RndIDSelected"] = _objectsService.RNDHome.IdSet; _objectsService.CLists.UpdateEmployee();
                }
                else _objectsService.RNDHome.IdSet = iOldSet;
            }
           return new JsonResult(true);

        }


        public DateTime SetDateTimeField(string sField)
        {
            DateTime dt = new DateTime(1900, 1, 1);
            if (!(_objectsService.RNDHome.dr[sField] == DBNull.Value)) dt = (DateTime)_objectsService.RNDHome.dr[sField];
            //          dt = (DateTime) null;
            return dt;
        }

        public IActionResult OnPostGExport_Click()
        {

            DataTable dt2 = new DataTable(); string sMsg, sFile;
            int icol, irow;
          //  Mouse.OverrideCursor = Cursors.Wait;

            string sqlBase = "  Select DateDSCreated, RN.[Study Name], RN.[Product ID], R3.[Product Description], R5.sLocation, R1.Employees as 'Operator', R2.Employees, PropTestingComplete, AgedTestingComplete, R4.[ID] ,R4.[IDDataset] ,R4.[FormNo] ,[NCOIndex] ,[IsoPbw] ,[DensT1],[DensT2] ,[DensT3] ,[DensT4] ,[DensT5] ,[DensL1] ,[DensL2],[DensW1] ,[DensW2],[DensAvgT],[DensAvgL],[DensAvgW],[DensMass],[Density],[CompStr1],[CompStr2],[CompStr3],[CompStr4],[CompStr],[CellDiaTop],[CellStDevTop],[CellCountTop],[CellDiaSide],[CellStDevSide],[CellCountSide],[CellDia],[CellDiaIsotropy] ,[CellCount],[ClosedCellPer1] ,[ClosedCellPer2],[ClosedCellPer3] ,[ClosedCellPer] ,[HotPlateInitMass],[HotPlateFinalMass] ,[HotPlateInitH1] ,[HotPlateInitH2],[HotPlateInitH3] ,[HotPlateInitH4],[HotPlateInitH5],[HotPlateInitH] ,[HotPlateFinalH1] ,[HotPlateFinalH2] ,[HotPlateFinalH3],[HotPlateFinalH4],[HotPlateFinalH5],[HotPlateFinalH],[HotPlateRetainMass],[HotPlateRetainThick],[ReactMixingTime] ,[React15PTime],[React50PTime] ,[React80PTime],[ReactCupEdgeTime],[React98PTime],[ReactMaxTempTime],[ReactMaxTemp],[ReactMaxHeight],[ReactSampleMass],[PhotoPirPur],[PhotoIso],[PhotoCarbo],[PhotoTrimer],[sFileFTIR],[sFileTGA],[sFileFoamat],R4.[Product Code],[K10D25FInit],[K10D40FInit] ,[K10D75FInit],[K10D110FInit],[K10D25FFinal],[K10D40FFinal],[K10D75FFinal] ,[K10D110FFinal] ,[K90D25FInit],[K90D40FInit],[K90D75FInit] ,[K90D110FInit] ,[K90D25FFinal] ,[K90D40FFinal] ,[K90D75FFinal] ,[K90D110FFinal] ,[K180D25FInit],[K180D40FInit] ,[K180D75FInit] ,[K180D110FInit] ,[K180D25FFinal],[K180D40FFinal] ,[K180D75FFinal],[K180D110FFinal] ,[R10D25FInit],[R10D40FInit],[R10D75FInit] ,[R10D110FInit],[R10D25FFinal] ,[R10D40FFinal] ,[R10D75FFinal] ,[R10D110FFinal],[R90D25FInit],[R90D40FInit] ,[R90D75FInit] ,[R90D110FInit] ,[R90D25FFinal],[R90D40FFinal] ,[R90D75FFinal],[R90D110FFinal],[R180D25FInit],[R180D40FInit] ,[R180D75FInit],[R180D110FInit],[R180D25FFinal],[R180D40FFinal],[R180D75FFinal] ,[R180D110FFinal] ,[sNote],' ' as 'Empty1', '' as 'Empty2','' as 'Empty3',R4.NCOIndex, RN.IsoMats, RN.sIsoMatsNCO, RN.POMats , RN.sPOMatsOH as 'OH#s',R4.POPbws as PBW from dbo.RNDDatasets as RN  Left JOIN Roster AS R1 ON RN.Operator = R1.ID  Left JOIN Roster AS R2 ON RN.Chemist = R2.ID   Left Join[Product Matrix] as R3 on RN.[Product ID] = [Product Code]    right join[RNDFormulations] as R4 on RN.ID = R4.IDDataset   Left join dbo.tblLocations as R5 on RN.Location = R5.ID ";

            string sql = GetSearchCriteria();

            //Final sql string
            if (!string.IsNullOrEmpty(sql)) sql = sqlBase + " Where " + sql; else sql = sqlBase;

            sql += " order by RN.[DateDSCreated] DESC";

            try
            {
                _objectsService.RNDHome.da = new SqlDataAdapter(sql, _objectsService.Cbfile.conAZ);
                dt2.Clear();
                int itmp = _objectsService.RNDHome.da.Fill(dt2);
                if (itmp < 1)
                {
                    sMsg = ("No R&D Dataset was found to meet the search criteria.  Check the network connection and/or relax the search criteria");
                  //  MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Information);
                    //                   return false;
                }
            }
            catch (SqlException ex)
            {
                sMsg = "Error in searching the R&D datasets for exporting";
             //   MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Stop);
                System.Diagnostics.Trace.TraceError(sMsg + "\n\n" + ex.Message);
                //  Mouse.OverrideCursor = null;
                // CTelClient.TelException(ex, sMsg);
                return new JsonResult(new { Message = true });
            }

            /*
                        SaveFileDialog FileDialog1 = new SaveFileDialog
                        {
                            Title = "Intugent PI - Export Mfg data to a tab limited data file ",
                            Filter = "Tab delimited File (*.txt) |*.txt|All Files (*.*)|*.*",
                        };

                        if (FileDialog1.ShowDialog() == false) { Mouse.OverrideCursor = null; return; }

                        Mouse.OverrideCursor = Cursors.Wait;

                        sFile = FileDialog1.FileName;

                        if (File.Exists(sFile))
                        {
                            sMsg = "File '" + sFile + "' already existis.  \n\nOverwrite it?";
                            if (MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) { Mouse.OverrideCursor = null; return; }
                        }

                        try
                        {
                            string[] sLines = new string[dt2.Rows.Count + 3];

                            sLines[0] = dt2.Columns[0].ColumnName.ToString();
                            for (icol = 1; icol < dt2.Columns.Count; icol++) sLines[0] += "\t" + dt2.Columns[icol].ColumnName.ToString();
                            for (irow = 0; irow < dt2.Rows.Count; irow++)
                            {
                                sLines[irow + 1] = (dt2.Rows[irow][0]).ToString();
                                for (icol = 1; icol < dt2.Columns.Count; icol++) sLines[irow + 1] += "\t" + (dt2.Rows[irow][icol]).ToString();
                            }

                            File.WriteAllLinesAsync(sFile, sLines);
                        }
                        catch (Exception ex) 
                        {sMsg = "Error in writting to the datasets to the export file"; MessageBox.Show(sMsg, Cbfile.sAppName); CTelClient.TelException(ex, sMsg);
                        }


                        CStatusBar.SetText("Exported data to '" + sFile + '"');
            */

            try
            {
                string[] sLines = new string[dt2.Rows.Count + 3];

                sLines[0] = dt2.Columns[0].ColumnName.ToString();
                for (icol = 1; icol < dt2.Columns.Count; icol++) sLines[0] += "\t" + dt2.Columns[icol].ColumnName.ToString();
                for (irow = 0; irow < dt2.Rows.Count; irow++)
                {
                    sLines[irow + 1] = (dt2.Rows[irow][0]).ToString();
                    for (icol = 1; icol < dt2.Columns.Count; icol++) sLines[irow + 1] += "\t" + (dt2.Rows[irow][icol]).ToString();
                }
              //  _objectsService.Cbfile.SendEmail(sLines);

            }
            catch (Exception ex)
            {
             //   sMsg = "Error in writting to the datasets to the export file"; MessageBox.Show(sMsg, Cbfile.sAppName); CTelClient.TelException(ex, sMsg);
            }
            // Mouse.OverrideCursor = null;

            return new JsonResult(new { Message = true });
        }

        public IActionResult  OnPostGCopy_Click()
        {
            DataTable dt2 = new DataTable(); string sMsg, sFile;
            int icol, irow;
          //  Mouse.OverrideCursor = Cursors.Wait;

            string sqlBase = "  Select DateDSCreated, RN.[Study Name], RN.[Product ID], R3.[Product Description], R5.sLocation, R1.Employees as 'Operator', R2.Employees, PropTestingComplete, AgedTestingComplete, R4.[ID] ,R4.[IDDataset] ,R4.[FormNo] ,[NCOIndex] ,[IsoPbw] ,[DensT1],[DensT2] ,[DensT3] ,[DensT4] ,[DensT5] ,[DensL1] ,[DensL2],[DensW1] ,[DensW2],[DensAvgT],[DensAvgL],[DensAvgW],[DensMass],[Density],[CompStr1],[CompStr2],[CompStr3],[CompStr4],[CompStr],[CellDiaTop],[CellStDevTop],[CellCountTop],[CellDiaSide],[CellStDevSide],[CellCountSide],[CellDia],[CellDiaIsotropy] ,[CellCount],[ClosedCellPer1] ,[ClosedCellPer2],[ClosedCellPer3] ,[ClosedCellPer] ,[HotPlateInitMass],[HotPlateFinalMass] ,[HotPlateInitH1] ,[HotPlateInitH2],[HotPlateInitH3] ,[HotPlateInitH4],[HotPlateInitH5],[HotPlateInitH] ,[HotPlateFinalH1] ,[HotPlateFinalH2] ,[HotPlateFinalH3],[HotPlateFinalH4],[HotPlateFinalH5],[HotPlateFinalH],[HotPlateRetainMass],[HotPlateRetainThick],[ReactMixingTime] ,[React15PTime],[React50PTime] ,[React80PTime],[ReactCupEdgeTime],[React98PTime],[ReactMaxTempTime],[ReactMaxTemp],[ReactMaxHeight],[ReactSampleMass],[PhotoPirPur],[PhotoIso],[PhotoCarbo],[PhotoTrimer],[sFileFTIR],[sFileTGA],[sFileFoamat],R4.[Product Code],[K10D25FInit],[K10D40FInit] ,[K10D75FInit],[K10D110FInit],[K10D25FFinal],[K10D40FFinal],[K10D75FFinal] ,[K10D110FFinal] ,[K90D25FInit],[K90D40FInit],[K90D75FInit] ,[K90D110FInit] ,[K90D25FFinal] ,[K90D40FFinal] ,[K90D75FFinal] ,[K90D110FFinal] ,[K180D25FInit],[K180D40FInit] ,[K180D75FInit] ,[K180D110FInit] ,[K180D25FFinal],[K180D40FFinal] ,[K180D75FFinal],[K180D110FFinal] ,[R10D25FInit],[R10D40FInit],[R10D75FInit] ,[R10D110FInit],[R10D25FFinal] ,[R10D40FFinal] ,[R10D75FFinal] ,[R10D110FFinal],[R90D25FInit],[R90D40FInit] ,[R90D75FInit] ,[R90D110FInit] ,[R90D25FFinal],[R90D40FFinal] ,[R90D75FFinal],[R90D110FFinal],[R180D25FInit],[R180D40FInit] ,[R180D75FInit],[R180D110FInit],[R180D25FFinal],[R180D40FFinal],[R180D75FFinal] ,[R180D110FFinal] ,[sNote],' ' as 'Empty1', '' as 'Empty2','' as 'Empty3',R4.NCOIndex, RN.IsoMats, RN.sIsoMatsNCO, RN.POMats , RN.sPOMatsOH as 'OH#s',R4.POPbws as PBW from dbo.RNDDatasets as RN  Left JOIN Roster AS R1 ON RN.Operator = R1.ID  Left JOIN Roster AS R2 ON RN.Chemist = R2.ID   Left Join[Product Matrix] as R3 on RN.[Product ID] = [Product Code]    right join[RNDFormulations] as R4 on RN.ID = R4.IDDataset   Left join dbo.tblLocations as R5 on RN.Location = R5.ID ";

            string sql = GetSearchCriteria();

            //Final sql string
            if (!string.IsNullOrEmpty(sql)) sql = sqlBase + " Where " + sql; else sql = sqlBase;

            sql += " order by RN.[DateDSCreated] DESC";

            try
            {
                _objectsService.RNDHome.da = new SqlDataAdapter(sql, _objectsService.Cbfile.conAZ);
                dt2.Clear();
                int itmp = _objectsService.RNDHome.da.Fill(dt2);
                if (itmp < 1)
                {
                    sMsg = ("No R&D Dataset was found to meet the search criteria.  Check the network connection and/or relax the search criteria");
                  //  MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Information);
                    //                   return false;
                }
            }
            catch (SqlException ex)
            {
                sMsg = "Error in searching the R&D datasets for exporting";
              //  MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.OK, MessageBoxImage.Stop);
                System.Diagnostics.Trace.TraceError(sMsg + "\n\n" + ex.Message);
              //  Mouse.OverrideCursor = null;
               // CTelClient.TelException(ex, sMsg);
                return null;
            }

            try
            {
                var sData = new StringBuilder();
                sData.Append(dt2.Columns[0].ColumnName.ToString());
                for (icol = 1; icol < dt2.Columns.Count; icol++) sData.Append("\t" + dt2.Columns[icol].ColumnName.ToString());
                for (irow = 0; irow < dt2.Rows.Count; irow++)
                {
                    sData.Append("\n" + (dt2.Rows[irow][0]).ToString());
                    for (icol = 1; icol < dt2.Columns.Count; icol++) sData.Append("\t" + (dt2.Rows[irow][icol]).ToString());
                }

                //               Cbfile.SendEmail(sLines);
               // Clipboard.SetText(sData.ToString());

            }
            catch (Exception ex)
            {
                //sMsg = "Error in writting datasets to the file"; MessageBox.Show(sMsg, Cbfile.sAppName); CTelClient.TelException(ex, sMsg);
               // Mouse.OverrideCursor = null;

            }



            /*
                        SaveFileDialog FileDialog1 = new SaveFileDialog
                        {
                            Title = "Intugent PI - Export Mfg data to a tab limited data file ",
                            Filter = "Tab delimited File (*.txt) |*.txt|All Files (*.*)|*.*",
                        };

                        if (FileDialog1.ShowDialog() == false) { Mouse.OverrideCursor = null; return; }

                        Mouse.OverrideCursor = Cursors.Wait;

                        sFile = FileDialog1.FileName;

                        if (File.Exists(sFile))
                        {
                            sMsg = "File '" + sFile + "' already existis.  \n\nOverwrite it?";
                            if (MessageBox.Show(sMsg, Cbfile.sAppName, MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) { Mouse.OverrideCursor = null; return; }
                        }

                        try
                        {
                            string[] sLines = new string[dt2.Rows.Count + 3];

                            sLines[0] = dt2.Columns[0].ColumnName.ToString();
                            for (icol = 1; icol < dt2.Columns.Count; icol++) sLines[0] += "\t" + dt2.Columns[icol].ColumnName.ToString();
                            for (irow = 0; irow < dt2.Rows.Count; irow++)
                            {
                                sLines[irow + 1] = (dt2.Rows[irow][0]).ToString();
                                for (icol = 1; icol < dt2.Columns.Count; icol++) sLines[irow + 1] += "\t" + (dt2.Rows[irow][icol]).ToString();
                            }

                            File.WriteAllLinesAsync(sFile, sLines);
                        }
                        catch (Exception ex) 
                        {sMsg = "Error in writting to the datasets to the export file"; MessageBox.Show(sMsg, Cbfile.sAppName); CTelClient.TelException(ex, sMsg);
                        }


                        CStatusBar.SetText("Exported data to '" + sFile + '"');
            

            try
            {
                string[] sLines = new string[dt2.Rows.Count + 3];

                sLines[0] = dt2.Columns[0].ColumnName.ToString();
                for (icol = 1; icol < dt2.Columns.Count; icol++) sLines[0] += "\t" + dt2.Columns[icol].ColumnName.ToString();
                for (irow = 0; irow < dt2.Rows.Count; irow++)
                {
                    sLines[irow + 1] = (dt2.Rows[irow][0]).ToString();
                    for (icol = 1; icol < dt2.Columns.Count; icol++) sLines[irow + 1] += "\t" + (dt2.Rows[irow][icol]).ToString();
                }
                Cbfile.SendEmail(sLines);

            }
            catch (Exception ex)
            {
                sMsg = "Error in writting to the datasets to the export file"; MessageBox.Show(sMsg, Cbfile.sAppName); CTelClient.TelException(ex, sMsg);
            }

            */
          //  Mouse.OverrideCursor = null;
            return new JsonResult(new {Message= true});
        }

        public string GetSearchCriteria()
        {
            DateTime dateTime;
            string sql = string.Empty, sql1 = string.Empty;

            _objectsService.RNDHome.sParamValue1 = string.Empty;
            if (_objectsService.CLists.drEmployee["RndDate1"] != DBNull.Value)
            {
                dateTime = ((DateTime)_objectsService.CLists.drEmployee["RndDate1"]).AddDays(1);
                sql1 = "DateDSCreated < '" + dateTime.ToString() + "'";
                if (sql == string.Empty) sql = sql1; else sql = sql + " And " + sql1;
            }

            if (_objectsService.CLists.drEmployee["RndDate2"] != DBNull.Value)
            {
                dateTime = ((DateTime)_objectsService.CLists.drEmployee["RndDate2"]);
                sql1 = "DateDSCreated >= '" + dateTime.ToString() + "'";
                if (sql == string.Empty) sql = sql1; else sql = sql + " And " + sql1;
            }

            sql1 = string.Empty;
            if (_objectsService.CLists.drEmployee["RndIDStudyType"] != DBNull.Value)
                if ((int)_objectsService.CLists.drEmployee["RndIDStudyType"] != 59) sql1 = " RN.[Study Type] = " + ((int)_objectsService.CLists.drEmployee["RndIDStudyType"]).ToString();
            if (sql1 != string.Empty) { if (sql == string.Empty) sql = sql1; else sql = sql + " And " + sql1; }

            sql1 = string.Empty;
            if (_objectsService.CLists.drEmployee["RndIDTestingStatus"] != DBNull.Value)
            {
                if ((int)_objectsService.CLists.drEmployee["RndIDTestingStatus"] == 52) sql1 = " RN.[PropTestingComplete] = 'false' and (RN.[Abandoned] is null or RN.[Abandoned] ='false' ) ";   //Testing in Progress
                else if ((int)_objectsService.CLists.drEmployee["RndIDTestingStatus"] == 53) sql1 = " RN.[PropTestingComplete] = 'true' and (RN.[Abandoned] is null or RN.[Abandoned] ='false' ) "; //Testing complete
                else if ((int)_objectsService.CLists.drEmployee["RndIDTestingStatus"] == 54) sql1 = " RN.[AgedTestingComplete] = 'true' and (RN.[Abandoned] is null or RN.[Abandoned] ='false' ) "; //Aged Testing Complete
                //                else if ((int)_objectsService.CLists.drEmployee["RndIDTestingStatus"] == 55) sql1 = " RN.[AgedTestingComplete] = true" 
                else if ((int)_objectsService.CLists.drEmployee["RndIDTestingStatus"] == 56) sql1 = " (RN.[Abandoned] is null or RN.[Abandoned] ='false' ) "; //All Active Dataset
                else if ((int)_objectsService.CLists.drEmployee["RndIDTestingStatus"] == 64) sql1 = " RN.[Abandoned]  ='true' "; //Abandoned datasets

            }
            if (sql1 != string.Empty) { if (sql == string.Empty) sql = sql1; else sql = sql + " And " + sql1; }

            sql1 = string.Empty;
            if (_objectsService.CLists.drEmployee["Rnd Product Code"] != DBNull.Value)
            {
                if ((string)_objectsService.CLists.drEmployee["Rnd Product Code"] != "All Products") sql1 = " RN.[Product ID] = '" + (string)_objectsService.CLists.drEmployee["Rnd Product Code"] + "' ";
            }
            if (sql1 != string.Empty) { if (sql == string.Empty) sql = sql1; else sql = sql + " And " + sql1; }

            sql1 = _objectsService.RNDHome.sParamValue1 = string.Empty;
            if (_objectsService.CLists.drEmployee["RNDNameSearch"] != DBNull.Value)
            {
                _objectsService.RNDHome.sParamValue1 = "%" + _objectsService.CLists.drEmployee["RNDNameSearch"].ToString() + "%";
                sql1 = "RN.[Study Name] Like @sParam1";
            }
            if (sql1 != string.Empty) { if (sql == string.Empty) sql = sql1; else sql = sql + " And " + sql1; }

            return sql;

        }

        public void SetTimeTextField(string sControl, string sField)
        {


            switch (sControl)
            {
                case "gTestTime":
                    //                    if (dr[sField] == DBNull.Value) gTestTime.CustomFormat = sTimeNullFormat;
                    //                   else { gTestTime.CustomFormat = sTimeFormat; gTestTime.Value = (DateTime)dr[sField]; }
                    break;
            }

        }

        #endregion
        public void Startup()
        {
            int idTemp;

            if (_objectsService.CLists.drEmployee["RndDate1"] == DBNull.Value) (gRndDate1) = null; else (gRndDate1) = (DateTime)_objectsService.CLists.drEmployee["RndDate1"];
            if (_objectsService.CLists.drEmployee["RndDate2"] == DBNull.Value) (gRndDate2) = null; else (gRndDate2) = (DateTime)_objectsService.CLists.drEmployee["RndDate2"];
            if (_objectsService.CLists.drEmployee["RNDNameSearch"] == DBNull.Value) gRNDNameSearch = null; else gRNDNameSearch = (string)_objectsService.CLists.drEmployee["RNDNameSearch"];
            if (_objectsService.CLists.drEmployee["Rnd Product Code"] == DBNull.Value) gProd1SelectedValue = _objectsService.CDefualts.sProdRNDAll; else gProd1SelectedValue = (string)_objectsService.CLists.drEmployee["Rnd Product Code"];
            if (_objectsService.CLists.drEmployee["RndIDTestingStatus"] == DBNull.Value) gTestStat1SelectedValue = _objectsService.RNDHome.CDefualts.iRNDTestingStat; else gTestStat1SelectedValue = (int)_objectsService.CLists.drEmployee["RndIDTestingStatus"];
            if (_objectsService.CLists.drEmployee["RndIDStudyType"] == DBNull.Value) gStudyTypeSelectedValue = _objectsService.RNDHome.CDefualts.iMfgRunType; else gStudyTypeSelectedValue = (int)_objectsService.CLists.drEmployee["RndIDStudyType"];

            if (_objectsService.CLists.drEmployee["RndSql"] != DBNull.Value) _objectsService.RNDHome.sqlSearchDS = (string)_objectsService.CLists.drEmployee["RndSql"];
            if (SearchRNDDB() && _objectsService.RNDHome.dt.Rows.Count > 0)
            {
                _objectsService.RNDHome.indSet = 0; _objectsService.RNDHome.IdSet = (int)_objectsService.RNDHome.dt.Rows[0]["ID"];
                if (_objectsService.CLists.drEmployee["RndIDSelected"] != DBNull.Value)
                {
                    idTemp = (int)_objectsService.CLists.drEmployee["RndIDSelected"];
                    for (int i = 0; i < _objectsService.RNDHome.dt.Rows.Count; i++) if ((int)_objectsService.RNDHome.dt.Rows[i]["ID"] == idTemp) { _objectsService.RNDHome.indSet = i; _objectsService.RNDHome.IdSet = (int)_objectsService.RNDHome.dt.Rows[i]["ID"]; }
                }
                gRNDSearchSelectedIndex = _objectsService.RNDHome.indSet;
                gRNDSearch = _objectsService.RNDHome.dt.DefaultView;
                _objectsService.RNDHome.EnableRNDPages(true);

            }
            else { _objectsService.RNDHome.EnableRNDPages(false); gRNDSearch = _objectsService.RNDHome.dt.DefaultView; }
        }


    }

}
