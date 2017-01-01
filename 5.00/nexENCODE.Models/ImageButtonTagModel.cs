using nexENCODE.Enum;
namespace nexENCODE.Models {
    /// <summary>
    /// Image Button Tag Model
    /// </summary>
    public class ImageButtonTagModel {
        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Button Type
        /// </summary>
        public ButtonTypes ButtonType { get; set; }
        /// <summary>
        /// File Name 1
        /// </summary>
        public string FileName1 { get; set; }
        /// <summary>
        /// File Name 2
        /// </summary>
        public string FileName2 { get; set; }
        /// <summary>
        /// File Name 3
        /// </summary>
        public string FileName3 { get; set; }
    }
}