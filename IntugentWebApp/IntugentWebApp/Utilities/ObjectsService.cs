using IntugentClassLbrary.Classes;
using IntugentClassLibrary.Pages.Mfg;

namespace IntugentWebApp.Utilities
{
    public class ObjectsService
    {
        public CDefualts CDefualts { get; set; }
        public CLists CLists { get; set; }
        public Cbfile Cbfile { get; set; }
        public MfgHome MfgHome { get; set; }
        public MfgInProcess? MfgInProcess { get; set; }
    }
}
