using IntugentClassLibrary.Classes;
using IntugentWebApp.Utilities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data;

namespace IntugentWebApp.Pages.Analysis
{
    [BindProperties]
    public class Analysis1Model : PageModel
    {
        string sSqlQuery;
        private readonly ObjectsService _objectsService;
        public DataView gProd1;
        public int gProd1SelectedIndex;
        public string gProd1SelectedValue;
        public DataView gMfgSite;
        public int gMfgSiteSelectedIndex;
        public string gMfgSiteSelectedValue;
        public DataView gProp1;
        public DataView gCorr;
        public string gProp1SelectedValue;
        public DateTime? gMfgDate1;
        public DateTime? gMfgDate2;
        public List<double> XA;
        public List<double> YA;
        public List<double> XAvg1;
        public List<double> YAvg1;
        public List<double> YUCL1;
        public List<double> YLCL1;

        public Analysis1Model(ObjectsService objectsService)
        {
            _objectsService = objectsService;

            _objectsService.CAnalysisData1.GetLists();
            gProd1 = _objectsService.CAnalysisData1.dtGlobalProducts.DefaultView;
            gMfgSite = _objectsService.CAnalysisData1.dtLocations.DefaultView;
            gProp1= _objectsService.CAnalysisData1.dtProps.DefaultView;
            gProp1SelectedValue = "FG-Compressive Str Avg6";
            gMfgDate2 = DateTime.Now;
            gMfgDate1 = DateTime.Now.AddYears(-1);

            //            gLoc1.ItemsSource = CAnalysisData.dtLocations.DefaultView;
        }

        public void OnGet()
        {
            if (gProp1SelectedValue == null) gProp1SelectedValue = "FG-Compressive Str Avg6";
            if (gProd1SelectedValue == null) gProd1SelectedIndex = 0;
            if (gMfgSiteSelectedValue == null) gMfgSiteSelectedIndex = 0;
            _objectsService.CAnalysisData1.GetData(GetSearchCriteria());
            SetView();
        }
        public IActionResult OnPostGSearchDB_Click()
        {
            if (gProp1SelectedValue == null) gProp1SelectedValue = "FG-Compressive Str Avg6";
            if (gProd1SelectedValue == null) gProd1SelectedIndex = 0;
            if (gMfgSiteSelectedValue == null) gMfgSiteSelectedIndex = 0;
            _objectsService.CAnalysisData1.GetData(GetSearchCriteria());
            SetView();
            return new JsonResult(true);
        }
        public void SetView()
        {
            if (gProp1SelectedValue == null) return;
            _objectsService.CAnalysisData1.CorrMatrix((string)gProp1SelectedValue);
            _objectsService.CAnalysisData1.dtCorr.DefaultView.Sort = "AbsValue DESC";
            gCorr = _objectsService.CAnalysisData1.dtCorr.DefaultView;
            PlotSPC();
        }
        public IActionResult OnPostGProp1_Changed()
        {
            SetView();
            return new JsonResult(gProp1SelectedValue);

        }

        public void PlotSPC()
        {
            double[] X1, Y1;
            int ncount = 0;
            DataTable dt = _objectsService.CAnalysisData1.dtPropValues;
            double dtmp, dSig, davg;
            double[] Xavg = { 0.0, 0.0 }, Yavg = { 0.0, 0.0 }, YUCL = { 0.0, 0.0 }, YLCL = { 0.0, 0.0 };

            X1 = new double[dt.Rows.Count];
            Y1 = new double[dt.Rows.Count];

            if (gProp1SelectedValue == null)
            {
                XA=X1.ToList();
                YA=Y1.ToList();
                // gLine1.Plot(X1, Y1);
                // gPts1.Plot(X1, Y1);
                return;
            }
            string scn1 = "Green-Board Time Stamp";
            string scn2 = (string)gProp1SelectedValue;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][scn1] == DBNull.Value || dt.Rows[i][scn2] == DBNull.Value) continue;
                if (double.IsNaN((double)dt.Rows[i][scn1]) || double.IsNaN((double)dt.Rows[i][scn2])) continue;
                X1[ncount] = (double)dt.Rows[i][scn1];
                Y1[ncount] = (double)dt.Rows[i][scn2];
                ncount++;
            }
            if (ncount < 1)
            {
                XA = X1.ToList();
                YA = Y1.ToList();
                // gLine1.Plot(X1, Y1);
                //gPts1.Plot(X1, Y1);
                return;
            }
            Array.Resize(ref X1, ncount);
            Array.Resize(ref Y1, ncount);
            //            double[] c = { ncount };
            //           double[] d = new double[ncount];

            davg = dSig = 0.0;
            for (int i = 0; i < Y1.Length; i++) davg += Y1[i];
            if (Y1.Length > 0) davg = davg / ((double)ncount);
            for (int i = 0; i < Y1.Length; i++) dSig += (Y1[i] - davg) * (Y1[i] - davg);
            if (Y1.Length > 0) dSig = Math.Sqrt(dSig / ((double)ncount));

            Xavg[0] = X1[0]; Xavg[1] = X1[X1.Length - 1];
            Yavg[0] = Yavg[1] = davg;
            YUCL[0] = YUCL[1] = davg + 2.0 * dSig;
            YLCL[0] = YLCL[1] = davg - 2.0 * dSig;
            XA = X1.ToList();
            YA = Y1.ToList();
            XAvg1 = Xavg.ToList();
            YAvg1 = Yavg.ToList();
            YLCL1 = YLCL.ToList();
            YUCL1 = YUCL.ToList();
            // gAvg.Plot(Xavg, Yavg);
            // gUCL.Plot(Xavg, YUCL);
            // gLCL.Plot(Xavg, YLCL);
            // gLine1.Plot(X1, Y1);
            // gPts1.Plot(X1, Y1);
            // //            gPts1.PlotColorSize(X1, Y1, c, c);
            // gSPC.BottomTitle = scn1 + " - " + _objectsService.CAnalysisData1.refDate.ToString();
            // gSPC.LeftTitle = scn2;

        }
        public string GetSearchCriteria()
        {
            string sql = string.Empty, sql1 = string.Empty;
            DateTime dateTime1;
            //Location

            sql = sql1 = string.Empty;
            if (gMfgSiteSelectedIndex > 0) sql = "sLocation = '" + gMfgSiteSelectedValue.ToString() + "'";


            sql1 = string.Empty;
            if (gProd1SelectedIndex > 0)
            {
                sql1 = "[Product Code Global] = '" + gProd1SelectedValue.ToString() + "'";
                if (sql == string.Empty) sql = sql1; else sql = sql + " And " + sql1;
            }

            //Dates
            if (gMfgDate2 != null)
            {
                dateTime1 = ((DateTime)gMfgDate2).AddDays(1);
                sql1 = "[Test Date] < '" + dateTime1.ToString() + "'";
                if (sql == string.Empty) sql = sql1; else sql = sql + " And " + sql1;
            }

            if (gMfgDate1 != null)
            {
                dateTime1 = ((DateTime)gMfgDate1);
                sql1 = "[Test Date] >= '" + dateTime1.ToString() + "'";
                if (sql == string.Empty) sql = sql1; else sql = sql + " And " + sql1;
            }

            return sql;
        }

    }
}
