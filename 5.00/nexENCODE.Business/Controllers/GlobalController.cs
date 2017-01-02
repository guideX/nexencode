using nexENCODE.Models;
namespace nexENCODE.Business.Controllers {
    /// <summary>
    /// Global Controller
    /// </summary>
    public class GlobalController {
        /// <summary>
        /// Ini
        /// </summary>
        public IniFilesModel Ini { get; set; }
        /// <summary>
        /// Skins
        /// </summary>
        public SkinsController Skins { get; set; }
        /// <summary>
        /// Global Controller
        /// </summary>
        public GlobalController(string startupPath) {
            Skins = new SkinsController();
            Ini.Skins = startupPath + @"\data\config\skins.ini";
            Ini.Settings = startupPath + @"\data\config\settings.ini";
            Ini.WindowPos = startupPath + @"\data\config\windowpos.ini";
        }
    }
}