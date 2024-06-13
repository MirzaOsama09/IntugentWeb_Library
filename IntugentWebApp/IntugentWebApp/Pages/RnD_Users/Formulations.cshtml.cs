using IntugentWebApp.Controllers.RnD;
using IntugentWebApp.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace IntugentWebApp.Pages.RnD_Users
{
    public class FormulationsModel : PageModel
    {
        [BindProperty]
        public string? Id { get; set; }

        [BindProperty]
        public RNDSearchResult DataSet { get; set; }

        public async void OnGet()
        {
            try
            {
                Id = HttpContext.Session.GetString("selectedDatasetID");
                if (Id != null)
                {
                    DataSet = await FormulationsController.GetSearchResults(Id);
                }
            }
            catch (Exception ex)
            {
                TempData["ErrorOnServer"] = ex.Message;
            }
        }
    }
}
