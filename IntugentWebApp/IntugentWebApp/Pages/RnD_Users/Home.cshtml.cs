using IntugentClassLbrary.Classes;
using IntugentWebApp.Controllers.RnD;
using IntugentWebApp.Models;
using IntugentWebApp.Utilities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data.SqlClient;
using System.Data;
using IntugentClassLibrary.Utilities;

namespace IntugentWebApp.Pages.RnD_Users
{
    public class HomeModel : PageModel
    {
        [BindProperty]
        public string? selectedDatasetID { get; set; }

        public List<RNDSearchResult>? searchResults { get; set; }

        private readonly ObjectsService _objectsService;
        public HomeModel(ObjectsService objectsService)
        {
            _objectsService = objectsService;
        }
        public async void OnGet()
        {
            try
            {
                selectedDatasetID = _objectsService.CLists.drEmployee["MfgIDSelected"].ToString();
                searchResults = await RndHomeController.GetSearchResults();
            }
            catch (Exception ex)
            {
                TempData["ErrorOnServer"] = ex.Message;
            }
        }

        public IActionResult OnPostCheckboxSelected(string id)
        {

            _objectsService.CLists.drEmployee["MfgIDSelected"] = id;
            CLists_UpdateEmployee.UpdateEmployee(_objectsService.CLists);

            return new JsonResult(new { message = "Dataset selected: " + id });
        }

    }

}
