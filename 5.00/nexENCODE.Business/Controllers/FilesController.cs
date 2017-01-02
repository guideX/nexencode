using System.IO;
namespace nexENCODE.Business.Controllers {
    /// <summary>
    /// Files Controller
    /// </summary>
    public static class FilesController {
        /// <summary>
        /// Read Directory
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string ReadDirectory(string fileName) {
            var obj = Path.GetDirectoryName(fileName);
            if (obj.Right(1) != @"\") {
                return obj + @"\";
            } else {
                return obj;
            }
        }
    }
}