using IntugentWebApp.Utilities;
using IntugentClassLibrary.Classes;
using Microsoft.AspNetCore.Mvc;
using IntugentClassLibrary.Pages.Mfg;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data;

namespace IntugentWebApp.Pages.Mfg_Group
{
    public class MfgJetMixingModel : PageModel
    {
        public string gFRate_A { get; set; }
        public string gFRate_B { get; set; }
        public string gTemp_A { get; set; }
        public string gTemp_B { get; set; }
        public string gPres_A { get; set; }
        public string gPres_B { get; set; }
        public string gDens_A { get; set; }
        public string gDens_B { get; set; }
        public string gVisO_A { get; set; }
        public string gVisO_B { get; set; }
        public string gVisE_A { get; set; }
        public string gVisE_B { get; set; }
        public string gDiaMixChamb { get; set; }
        public string gDiaNoz_A { get; set; }
        public string gDiaNoz_B { get; set; }
        public string gPres_Max { get; set; }
        public string gPres_Min { get; set; }
        public string gReNo_Min { get; set; }
        public DataView gDetails { get; set; }
        public string gMsg { get; set; }
        public List<double> XA { get; set; }
        public List<double> YA { get; set; }
        public List<double> XB { get; set; }
        public List<double> YB { get; set; }

        private readonly ObjectsService _objectsService;

        public MfgJetMixingModel(ObjectsService objectsService)
        {
            _objectsService = objectsService;
        }
        public void OnGet()
        {
            SetView();
            SetFields();
        }
        public void SetView()
        {
            string sMsg = string.Empty;
            int ir = 0, i;

            _objectsService.MfgJetMixing.JetMix1.JetMixingSim();

            _objectsService.MfgJetMixing.dtDetails.Rows[ir]["Description"] = "Jet Volume Flow Rate [m3/s]";
            _objectsService.MfgJetMixing.dtDetails.Rows[ir]["Jet_A"] = _objectsService.MfgJetMixing.JetMix1.VolRate_A.ToString("0.000E00");
            _objectsService.MfgJetMixing.dtDetails.Rows[ir]["Jet_B"] = _objectsService.MfgJetMixing.JetMix1.VolRate_B.ToString("0.000E00");

            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 1]["Description"] = "Fluid Viscosity [Pa-s]";
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 1]["Jet_A"] = _objectsService.MfgJetMixing.JetMix1.Vis_A.ToString("0.000E00");
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 1]["Jet_B"] = _objectsService.MfgJetMixing.JetMix1.Vis_B.ToString("0.000E00");

            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 2]["Description"] = "Jet Diameter [m]";
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 2]["Jet_A"] = _objectsService.MfgJetMixing.JetMix1.DiaJet_A.ToString("0.000E00");
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 2]["Jet_B"] = _objectsService.MfgJetMixing.JetMix1.DiaJet_B.ToString("0.000E00");

            if (_objectsService.MfgJetMixing.JetMix1.DiaJet_A > _objectsService.MfgJetMixing.JetMix1.DiaNoz_A) sMsg = "Nozzle A must be bigger than " + (_objectsService.MfgJetMixing.JetMix1.DiaJet_A / _objectsService.MfgJetMixing.cUConv.ToSi_Dia).ToString("0.000E00") + " mm.  ";
            if (_objectsService.MfgJetMixing.JetMix1.DiaJet_B > _objectsService.MfgJetMixing.JetMix1.DiaNoz_B) sMsg += "Nozzle B must be bigger than " + (_objectsService.MfgJetMixing.JetMix1.DiaJet_B / _objectsService.MfgJetMixing.cUConv.ToSi_Dia).ToString("0.000E00") + " mm";
             gMsg = sMsg;

            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 3]["Description"] = "Jet Velocity [m/s]";
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 3]["Jet_A"] = _objectsService.MfgJetMixing.JetMix1.Vel_A.ToString("0.000E00");
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 3]["Jet_B"] = _objectsService.MfgJetMixing.JetMix1.Vel_B.ToString("0.000E00");


            ir += 1;
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 3]["Description"] = "Jet Reynold No.";
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 3]["Jet_A"] = _objectsService.MfgJetMixing.JetMix1.ReNo_A.ToString("0.000E00");
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 3]["Jet_B"] = _objectsService.MfgJetMixing.JetMix1.ReNo_B.ToString("0.000E00");

            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 4]["Description"] = "Jet Kinetic Energy [J/s]";
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 4]["Jet_A"] = _objectsService.MfgJetMixing.JetMix1.KE_A.ToString("0.000E00");
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 4]["Jet_B"] = _objectsService.MfgJetMixing.JetMix1.KE_B.ToString("0.000E00");

            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 5]["Description"] = "Impingement Point [r/D]";
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 5]["Jet_A"] = _objectsService.MfgJetMixing.JetMix1.Zeta.ToString("0.000E00");
            _objectsService.MfgJetMixing.dtDetails.Rows[ir + 5]["Jet_B"] = _objectsService.MfgJetMixing.JetMix1.Zeta.ToString("0.000E00");



            gDetails = _objectsService.MfgJetMixing.dtDetails.DefaultView;




            double dx = 1.0 / (double)(MfgJetMixing.nPts - 1);
            _objectsService.MfgJetMixing.XA[0] = -0.5; _objectsService.MfgJetMixing.XB[0] = 0.5;
            _objectsService.MfgJetMixing.YA[0] = _objectsService.MfgJetMixing.YB[0] = -0.5;

            //            _objectsService.MfgJetMixing.JetMix1.Zeta = 0.1;
            for (i = 1; i < MfgJetMixing.nPts; i++)
            {
                _objectsService.MfgJetMixing.YA[i] = _objectsService.MfgJetMixing.YA[i - 1] + dx;
                _objectsService.MfgJetMixing.YB[i] = _objectsService.MfgJetMixing.YA[i];
                _objectsService.MfgJetMixing.XA[i] = (_objectsService.MfgJetMixing.JetMix1.Zeta + 0.5) * (1 - 4.0 * _objectsService.MfgJetMixing.YA[i] * _objectsService.MfgJetMixing.YA[i]) - 0.5;
                _objectsService.MfgJetMixing.XB[i] = 0.5 - (0.5 - _objectsService.MfgJetMixing.JetMix1.Zeta) * (1 - 4.0 * _objectsService.MfgJetMixing.YA[i] * _objectsService.MfgJetMixing.YA[i]);
            }
            //gJetA.Plot(_objectsService.MfgJetMixing.XA, _objectsService.MfgJetMixing.YA);
            // gJetB.Plot(_objectsService.MfgJetMixing.XB, _objectsService.MfgJetMixing.YB);
            XA = _objectsService.MfgJetMixing.XA.ToList();
            YA = _objectsService.MfgJetMixing.YA.ToList();
            XB = _objectsService.MfgJetMixing.XB.ToList();
            YB = _objectsService.MfgJetMixing.YB.ToList();
        }
        public void SetFields()
        {
            gDetails = _objectsService.MfgJetMixing.dtDetails.DefaultView;
            gDiaMixChamb = (_objectsService.MfgJetMixing.JetMix1.DiaMixChamb / _objectsService.MfgJetMixing.cUConv.ToSi_Dia).ToString("0.000");
            gDiaNoz_A = (_objectsService.MfgJetMixing.JetMix1.DiaNoz_A / _objectsService.MfgJetMixing.cUConv.ToSi_Dia).ToString("0.000");
            gDiaNoz_B = (_objectsService.MfgJetMixing.JetMix1.DiaNoz_B / _objectsService.MfgJetMixing.cUConv.ToSi_Dia).ToString("0.000");
            gPres_Max = (_objectsService.MfgJetMixing.JetMix1.Pres_Max / _objectsService.MfgJetMixing.cUConv.ToSi_Pres).ToString("0");
            gPres_Min = (_objectsService.MfgJetMixing.JetMix1.Pres_Min / _objectsService.MfgJetMixing.cUConv.ToSi_Pres).ToString("0");
            gReNo_Min = _objectsService.MfgJetMixing.JetMix1.ReNo_Min.ToString("0.0");

            gFRate_A = (_objectsService.MfgJetMixing.JetMix1.FRate_A / _objectsService.MfgJetMixing.cUConv.ToSi_FRate).ToString("0.000");
            gFRate_B = (_objectsService.MfgJetMixing.JetMix1.FRate_B / _objectsService.MfgJetMixing.cUConv.ToSi_FRate).ToString("0.000");
            gTemp_A = (_objectsService.MfgJetMixing.JetMix1.Temp_A * 1.8 - 459.67).ToString("0.0");
            gTemp_B = (_objectsService.MfgJetMixing.JetMix1.Temp_B * 1.8 - 459.67).ToString("0.0");
            gPres_A = (_objectsService.MfgJetMixing.JetMix1.Pres_A / _objectsService.MfgJetMixing.cUConv.ToSi_Pres).ToString("0.0");
            gPres_B = (_objectsService.MfgJetMixing.JetMix1.Pres_B / _objectsService.MfgJetMixing.cUConv.ToSi_Pres).ToString("0.0");
            gDens_A = (_objectsService.MfgJetMixing.JetMix1.Dens_A / _objectsService.MfgJetMixing.cUConv.ToSi_Dens).ToString("0.00");
            gDens_B = (_objectsService.MfgJetMixing.JetMix1.Dens_B / _objectsService.MfgJetMixing.cUConv.ToSi_Dens).ToString("0.00");

            gVisO_A = (_objectsService.MfgJetMixing.JetMix1.VisO_A / _objectsService.MfgJetMixing.cUConv.ToSi_Vis).ToString("0.0");
            gVisO_B = (_objectsService.MfgJetMixing.JetMix1.VisO_B / _objectsService.MfgJetMixing.cUConv.ToSi_Vis).ToString("0.0");
            //            gVisN_A  = (_objectsService.MfgJetMixing.JetMix1.VisN_A ).ToString("0.000");
            //            gVisN_B  = (_objectsService.MfgJetMixing.JetMix1.VisN_B ).ToString("0.000");
            gVisE_A = (_objectsService.MfgJetMixing.JetMix1.VisE_A).ToString("0.0");
            gVisE_B = (_objectsService.MfgJetMixing.JetMix1.VisE_B).ToString("0.0");

            //            if (_objectsService.MfgJetMixing.JetMix1.bPlantSim) gJM_Plant.IsChecked = true;
            //            else gJM_Expert.IsChecked = true;



        }

        public IActionResult OnPostClick_gOptimize()
        {
            int nParam = 2, iter = 0, niter = 100;  // # of parameters and iterations
            double ftol = 0.0000001, fret = 100;
            double[] param = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            double[] pMax = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            double[] pMin = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            double[] paramOld = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            double tmpDia, tmpVel, tmpPMin, tmpPMin1, tmpPMin2, pi = Math.PI;


            //            SetParamLimits(pMax, pMin);
            //            _objectsService.MfgJetMixing.JetMix1.JetMixingSim();

            /*
             * 1. Calculate the velocity to with Jet Dia = Nozzle Dia and then calculate the pressure.  This is minimum pressure to avoid jet size greater than nozzle size.
             * 2. Calculate the diameter for the minimum Reynold number using Re# expression in terms of mass flow rate.  Then calculate velocity and pressure.  This is minimum pressure for the Reynold # to be above the min Reynold #.
             * 3. The minimum pressure is max of all the min pressures, but less or equal to the max pressure.
             * 
             */

            _objectsService.MfgJetMixing.JetMix1.JetMixingSim();  //Needed to calculate intermediate quants

            tmpVel = _objectsService.MfgJetMixing.JetMix1.VolRate_A * 4.0 / (pi * _objectsService.MfgJetMixing.JetMix1.DiaNoz_A * _objectsService.MfgJetMixing.JetMix1.DiaNoz_A);
            tmpPMin1 = 0.5 * _objectsService.MfgJetMixing.JetMix1.Dens_A * tmpVel * tmpVel;

            tmpDia = _objectsService.MfgJetMixing.JetMix1.FRate_A * 4.0 / (pi * _objectsService.MfgJetMixing.JetMix1.Vis_A * _objectsService.MfgJetMixing.JetMix1.ReNo_Min);
            tmpVel = _objectsService.MfgJetMixing.JetMix1.VolRate_A * 4.0 / (pi * tmpDia * tmpDia);
            tmpPMin2 = 0.5 * _objectsService.MfgJetMixing.JetMix1.Dens_A * tmpVel * tmpVel;

            tmpPMin = _objectsService.MfgJetMixing.JetMix1.Pres_Min;
            if (tmpPMin1 > tmpPMin) tmpPMin = tmpPMin1;
            if (tmpPMin2 > tmpPMin) tmpPMin = tmpPMin2;
            if (tmpPMin > _objectsService.MfgJetMixing.JetMix1.Pres_Max) tmpPMin = _objectsService.MfgJetMixing.JetMix1.Pres_Max;
            pMin[0] = tmpPMin;
            pMax[0] = _objectsService.MfgJetMixing.JetMix1.Pres_Max;
            /* Side B Calculations */
            tmpVel = _objectsService.MfgJetMixing.JetMix1.VolRate_B * 4.0 / (pi * _objectsService.MfgJetMixing.JetMix1.Dens_B * _objectsService.MfgJetMixing.JetMix1.DiaNoz_B * _objectsService.MfgJetMixing.JetMix1.DiaNoz_B);
            tmpPMin1 = 0.5 * _objectsService.MfgJetMixing.JetMix1.Dens_B * tmpVel * tmpVel;

            tmpDia = _objectsService.MfgJetMixing.JetMix1.FRate_B * 4.0 / (pi * _objectsService.MfgJetMixing.JetMix1.Vis_B * _objectsService.MfgJetMixing.JetMix1.ReNo_Min);
            tmpVel = _objectsService.MfgJetMixing.JetMix1.VolRate_B * 4.0 / (pi * tmpDia * tmpDia);
            tmpPMin2 = 0.5 * _objectsService.MfgJetMixing.JetMix1.Dens_B * tmpVel * tmpVel;

            tmpPMin = _objectsService.MfgJetMixing.JetMix1.Pres_Min;
            if (tmpPMin1 > tmpPMin) tmpPMin = tmpPMin1;
            if (tmpPMin2 > tmpPMin) tmpPMin = tmpPMin2;
            if (tmpPMin > _objectsService.MfgJetMixing.JetMix1.Pres_Max) tmpPMin = _objectsService.MfgJetMixing.JetMix1.Pres_Max;
            pMin[1] = tmpPMin;
            pMax[1] = _objectsService.MfgJetMixing.JetMix1.Pres_Max;

            /* Ensure that current values are within min and max pressures. */

            if (_objectsService.MfgJetMixing.JetMix1.Pres_A < pMin[0]) _objectsService.MfgJetMixing.JetMix1.Pres_A = pMin[0];
            if (_objectsService.MfgJetMixing.JetMix1.Pres_B < pMin[1]) _objectsService.MfgJetMixing.JetMix1.Pres_B = pMin[1];
            if (_objectsService.MfgJetMixing.JetMix1.Pres_A > pMax[0]) _objectsService.MfgJetMixing.JetMix1.Pres_A = pMax[0];
            if (_objectsService.MfgJetMixing.JetMix1.Pres_B > pMax[1]) _objectsService.MfgJetMixing.JetMix1.Pres_B = pMax[1];

            SetParameters(param, "read");
            SetParameters(param, "write");

            MyMathLib.Frprmn_1(param, pMax, pMin, nParam, ftol, ref iter, ref fret, PFunc, DFunc);
            SetFields();
            SetView();
            return new JsonResult(new { message = "success" });
        }

        public void SetParameters(double[] param, string action)
        {
            if (action == "read")
            {
                param[0] = _objectsService.MfgJetMixing.JetMix1.Pres_A;
                param[1] = _objectsService.MfgJetMixing.JetMix1.Pres_B;

            }
            else if (action == "write")
            {
                _objectsService.MfgJetMixing.JetMix1.Pres_A = param[0];
                _objectsService.MfgJetMixing.JetMix1.Pres_B = param[1];
            }
        }

        public void SetParamLimits(double[] pMax, double[] pMin)
        {
            pMax[0] = pMax[1] = _objectsService.MfgJetMixing.JetMix1.Pres_Max;
            pMin[0] = pMin[1] = _objectsService.MfgJetMixing.JetMix1.Pres_Min;

        }

        public void DFunc(int n, double[] p, double[] dfp)
        {
            int i, ip;
            double[] p1 = new double[n];
            double del = 0.001, delp, fp, fp1;

            fp = PFunc(p);
            for (ip = 0; ip < n; ip++)
            {
                for (i = 0; i < n; i++) p1[i] = p[i];
                if (p[ip] < 0.1 && p[ip] > -0.1)
                {
                    p1[ip] = del;
                    delp = del;
                }
                else
                {
                    p1[ip] = (1.0 + del) * p[ip];
                    delp = (p1[ip] - p[ip]) / Math.Abs(p[ip]);
                }
                fp1 = PFunc(p1);
                dfp[ip] = (fp1 - fp) / delp;
            }
        }

        public double PFunc(double[] p)
        {
            double dum = 0;
            SetParameters(p, "write");
            _objectsService.MfgJetMixing.JetMix1.JetMixingSim();
            dum = _objectsService.MfgJetMixing.JetMix1.Zeta * _objectsService.MfgJetMixing.JetMix1.Zeta;
            if (_objectsService.MfgJetMixing.JetMix1.ReNo_A < 100.0) dum += 0.25 * (1 - 0.01 * _objectsService.MfgJetMixing.JetMix1.ReNo_A);
            if (_objectsService.MfgJetMixing.JetMix1.ReNo_B < 100.0) dum += 0.25 * (1 - 0.01 * _objectsService.MfgJetMixing.JetMix1.ReNo_B);
            return dum;
        }

        //private void PreviewPositiveRealNumber(object sender, TextCompositionEventArgs e)
        //{
        //    var regex = new Regex("[^0-9 .]");
        //    e.Handled = regex.IsMatch(e );
        //}

        public IActionResult OnPostChange_gFRate_A()
        {
            if (Double.TryParse(gFRate_A, out double dum)) _objectsService.MfgJetMixing.JetMix1.FRate_A = dum;
            else gFRate_A= _objectsService.MfgJetMixing.JetMix1.FRate_A.ToString("0.000");
            return new JsonResult(new { message = "Success" });
        }

        public IActionResult OnPostLostFocus_Fields(string Name,string value)
        {
            string sFld = String.Empty, sMsg;
            string sName = Name;
            double dum=double.Parse(value);

            switch (sName)
            {
                case "gFRate_A":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.FRate_A = dum * _objectsService.MfgJetMixing.cUConv.ToSi_FRate;
                    else {
                       // MessageBox.Show("Improper Value. New Value is not accepted.");
                        gFRate_A  = (_objectsService.MfgJetMixing.JetMix1.FRate_A / _objectsService.MfgJetMixing.cUConv.ToSi_FRate).ToString("0.000"); }
                    break;
                case "gTemp_A":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.Temp_A = (dum + 459.67) / 1.8;
                    else { 
                        //MessageBox.Show("Improper Value. New Value is not accepted.");
                        gTemp_A  = (_objectsService.MfgJetMixing.JetMix1.Temp_A * 1.8 + 459.67).ToString("0.000"); }
                    break;
                case "gPres_A":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.Pres_A = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Pres;
                    else { 
                       // MessageBox.Show("Improper Value. New Value is not accepted."); 
                        gPres_A  = (_objectsService.MfgJetMixing.JetMix1.Pres_A / _objectsService.MfgJetMixing.cUConv.ToSi_Pres).ToString("0.000"); }
                    break;
                case "gFRate_B":
                    if(!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.FRate_B = dum * _objectsService.MfgJetMixing.cUConv.ToSi_FRate;
                    else { 
                       // MessageBox.Show("Improper Value. New Value is not accepted."); 
                        gFRate_B  = (_objectsService.MfgJetMixing.JetMix1.FRate_B / _objectsService.MfgJetMixing.cUConv.ToSi_FRate).ToString("0.000"); }
                    break;
                case "gTemp_B":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.Temp_B = (dum + 459.67) / 1.8;
                    else {
                       // MessageBox.Show("Improper Value. New Value is not accepted.");
                        gTemp_B  = (_objectsService.MfgJetMixing.JetMix1.Temp_B * 1.8 + 459.67).ToString("0.000"); }
                    break;
                case "gPres_B":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.Pres_B = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Pres;
                    else { 
                       // MessageBox.Show("Improper Value. New Value is not accepted.");
                        gPres_B  = (_objectsService.MfgJetMixing.JetMix1.Pres_B / _objectsService.MfgJetMixing.cUConv.ToSi_Pres).ToString("0.000"); }
                    break;

                case "gDens_A":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.Dens_A = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Dens;
                    else {
                        //MessageBox.Show("Improper Value. New Value is not accepted.");
                        gDens_A  = (_objectsService.MfgJetMixing.JetMix1.Dens_A / _objectsService.MfgJetMixing.cUConv.ToSi_Dens).ToString("0.000"); }
                    break;
                case "gVisO_A":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.VisO_A = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Vis;
                    else {
                      //  MessageBox.Show("Improper Value. New Value is not accepted.");
                        gVisO_A  = (_objectsService.MfgJetMixing.JetMix1.VisO_A / _objectsService.MfgJetMixing.cUConv.ToSi_Vis).ToString("0.000"); }
                    break;
                case "gVisE_A":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.VisE_A = dum;
                    else { 
                        //MessageBox.Show("Improper Value. New Value is not accepted.");
                        gVisE_A  = (_objectsService.MfgJetMixing.JetMix1.VisE_A).ToString("0.000"); }
                    break;
                case "gDens_B":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.Dens_B = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Dens;
                    else { 
                       // MessageBox.Show("Improper Value. New Value is not accepted."); 
                        gDens_B  = (_objectsService.MfgJetMixing.JetMix1.Dens_B / _objectsService.MfgJetMixing.cUConv.ToSi_Dens).ToString("0.000"); }
                    break;
                case "gVisO_B":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.VisO_B = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Vis;
                    else { 
                       // MessageBox.Show("Improper Value. New Value is not accepted."); 
                        gVisO_B  = (_objectsService.MfgJetMixing.JetMix1.VisO_B / _objectsService.MfgJetMixing.cUConv.ToSi_Vis).ToString("0.000"); }
                    break;
                case "gVisE_B":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.VisE_B = dum;
                    else {
                        //MessageBox.Show("Improper Value. New Value is not accepted.");
                        gVisE_B  = (_objectsService.MfgJetMixing.JetMix1.VisE_B).ToString("0.000"); }
                    break;

                case "gDiaMixChamb":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.DiaMixChamb = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Dia;
                    else {
                       // MessageBox.Show("Improper Value. New Value is not accepted.");
                        gDiaMixChamb  = (_objectsService.MfgJetMixing.JetMix1.DiaMixChamb / _objectsService.MfgJetMixing.cUConv.ToSi_Dia).ToString("0.000"); }
                    break;
                case "gDiaNoz_A":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.DiaNoz_A = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Dia;
                    else { 
                       // MessageBox.Show("Improper Value. New Value is not accepted."); 
                        gDiaNoz_A  = (_objectsService.MfgJetMixing.JetMix1.DiaNoz_A / _objectsService.MfgJetMixing.cUConv.ToSi_Dia).ToString("0.000"); }
                    break;
                case "gDiaNoz_B":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.DiaNoz_B = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Dia;
                    else { 
                        //MessageBox.Show("Improper Value. New Value is not accepted.");
                        gDiaNoz_B  = (_objectsService.MfgJetMixing.JetMix1.DiaNoz_B / _objectsService.MfgJetMixing.cUConv.ToSi_Dia).ToString("0.000"); }
                    break;

                case "gPres_Max":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.Pres_Max = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Pres;
                    else { 
                       // MessageBox.Show("Improper Value. New Value is not accepted.");
                        gPres_Max  = (_objectsService.MfgJetMixing.JetMix1.Pres_Max / _objectsService.MfgJetMixing.cUConv.ToSi_Pres).ToString("0.000"); }
                    break;
                case "gPres_Min":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.Pres_Min = dum * _objectsService.MfgJetMixing.cUConv.ToSi_Pres;
                    else { 
                       // MessageBox.Show("Improper Value. New Value is not accepted.");
                        gPres_Min  = (_objectsService.MfgJetMixing.JetMix1.Pres_Min / _objectsService.MfgJetMixing.cUConv.ToSi_Pres).ToString("0.000"); }
                    break;
                case "gReNo_Min":
                    if (!string.IsNullOrEmpty(value)) _objectsService.MfgJetMixing.JetMix1.ReNo_Min = dum;
                    else {
                      //  MessageBox.Show("Improper Value. New Value is not accepted."); 
                        gReNo_Min  = (_objectsService.MfgJetMixing.JetMix1.ReNo_Min).ToString("0.000"); }
                    break;
            }

            UpdateDataset();
            SetView();
            return new JsonResult(new { message = sName});
        }



        public IActionResult OnPostClick_gImportPourHeadData(string Name)
        {
            bool bNullValues = false; double dtemp;
            string sFld = String.Empty, sMsg;
           // Control ctrl = sender as Control;

            if (!_objectsService.MfgPlantsData.GetDataSet()) {
                sMsg = "Can not retrieve Pour Head Data"; 
               // MessageBox.Show(sMsg, Cbfile.sAppName); 
                return null; }
            DataRow dr = _objectsService.MfgPlantsData.dr;
            string sName = Name;

            switch (sName)
            {
                case "gImport1":
                    if (dr["MDI 1 Pour Head Flowrate"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["MDI 1 Pour Head Flowrate"]; gFRate_A  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.FRate_A = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_FRate; }
                    if (dr["Poly 1 Pour Head Flowrate"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["Poly 1 Pour Head Flowrate"]; gFRate_B  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.FRate_B = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_FRate; }

                    if (dr["MDI 1 Pour Head Temperature"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["MDI 1 Pour Head Temperature"]; gTemp_A  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Temp_A = (dtemp + 459.67) / 1.8; }
                    if (dr["Poly 1 Pour Head Temperature"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["Poly 1 Pour Head Temperature"]; gTemp_B  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Temp_B = (dtemp + 459.67) / 1.8; }

                    if (dr["MDI 1 Pour Head Pressure"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["MDI 1 Pour Head Pressure"]; gPres_A  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Pres_A = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_Pres; }
                    if (dr["Poly 1 Pour Head Pressure"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["Poly 1 Pour Head Pressure"]; gPres_B  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Pres_B = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_Pres; }
                    break;

                case "gImport2":
                    if (dr["MDI 2 Pour Head Flowrate"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["MDI 2 Pour Head Flowrate"]; gFRate_A  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.FRate_A = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_FRate; }
                    if (dr["Poly 2 Pour Head Flowrate"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["Poly 2 Pour Head Flowrate"]; gFRate_B  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.FRate_B = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_FRate; }

                    if (dr["MDI 2 Pour Head Temperature"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["MDI 2 Pour Head Temperature"]; gTemp_A  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Temp_A = (dtemp + 459.67) / 1.8; }
                    if (dr["Poly 2 Pour Head Temperature"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["Poly 2 Pour Head Temperature"]; gTemp_B  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Temp_B = (dtemp + 459.67) / 1.8; }

                    if (dr["MDI 2 Pour Head Pressure"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["MDI 2 Pour Head Pressure"]; gPres_A  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Pres_A = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_Pres; }
                    if (dr["Poly 2 Pour Head Pressure"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["Poly 2 Pour Head Pressure"]; gPres_B  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Pres_B = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_Pres; }
                    break;

                case "gImport3":
                    if (dr["MDI 3 Pour Head Flowrate"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["MDI 3 Pour Head Flowrate"]; gFRate_A  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.FRate_A = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_FRate; }
                    if (dr["Poly 3 Pour Head Flowrate"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["Poly 3 Pour Head Flowrate"]; gFRate_B  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.FRate_B = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_FRate; }

                    if (dr["MDI 3 Pour Head Temperature"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["MDI 3 Pour Head Temperature"]; gTemp_A  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Temp_A = (dtemp + 459.67) / 1.8; }
                    if (dr["Poly 3 Pour Head Temperature"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["Poly 3 Pour Head Temperature"]; gTemp_B  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Temp_B = (dtemp + 459.67) / 1.8; }

                    if (dr["MDI 3 Pour Head Pressure"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["MDI 3 Pour Head Pressure"]; gPres_A  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Pres_A = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_Pres; }
                    if (dr["Poly 3 Pour Head Pressure"] == DBNull.Value) bNullValues = true;
                    else { dtemp = (double)dr["Poly 3 Pour Head Pressure"]; gPres_B  = dtemp.ToString("0.000"); _objectsService.MfgJetMixing.JetMix1.Pres_B = dtemp * _objectsService.MfgJetMixing.cUConv.ToSi_Pres; }
                    break;
            }
            sMsg = "Some values were not available and were not changed";
            if (bNullValues)
                //MessageBox.Show(sMsg, Cbfile.sAppName);

            SetView();
            UpdateDataset();
            return new JsonResult(new { message = true });
        }



        public IActionResult OnPostClick_gReset()
        {
            _objectsService.MfgJetMixing.SetDefaultValues();
            SetFields();
            SetView();
            UpdateDataset();
            return new JsonResult(new { message = "Success" });
        }

        /*
                void Application_Closing(object sender, CancelEventArgs e)
                {
                    //Check to see if the present file has been given a name
                    try
                    {
                        BinaryFormatter formatter = new BinaryFormatter();
                        Stream stream = new FileStream(CAppParam.sFileName, FileMode.Create, FileAccess.Write);
                        // Disable the warning.
        #pragma warning disable SYSLIB0011
                        formatter.Serialize(stream, CAppParam.sAppName);
                        formatter.Serialize(stream, CAppParam.sVersion);
                        formatter.Serialize(stream, _objectsService.MfgJetMixing.JetMix1);

                        stream.Close();
                    }
                    catch { }

                }
        */
        public void UpdateDataset()
        {
           string js1 = System.Text.Json.JsonSerializer.Serialize(_objectsService.MfgJetMixing.JetMix1);
            _objectsService.CLists.drEmployee["sJetMix"] = js1;
            _objectsService.CLists.UpdateEmployee();
        }
    }
}
