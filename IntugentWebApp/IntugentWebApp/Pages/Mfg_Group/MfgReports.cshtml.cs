using IntugentWebApp.Utilities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data;

namespace IntugentWebApp.Pages.Mfg_Group
{
    public class MfgReportsModel : PageModel
    {
        public readonly ObjectsService _objectsService;
        public DateTime gReportDate {  get; set; }
        public DataView gReport {  get; set; }
        public MfgReportsModel(ObjectsService objectsService)
        {
            _objectsService = objectsService;

        }
        public void OnGet()
        {
            gReportDate = new DateTime(2023, 5, 10);
            //gReportDate= DateTime.Now;
            _objectsService.MfgReport.MfgReport(gReportDate);
            gReport = _objectsService.MfgReport.dt.DefaultView;
        }

        public  IActionResult OnPostReportDate_Changed(DateTime value)
        {
            if (!_objectsService.MfgReport.binit) _objectsService.MfgReport.MfgReport(value);
            _objectsService.MfgReport.binit = false;
            gReport = _objectsService.MfgReport.dt.DefaultView;
            return new JsonResult(new { message = value});
        }
    }
}
