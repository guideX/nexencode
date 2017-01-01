namespace nexENCODE.Models.Skin {
    /// <summary>
    /// Skins Model
    /// </summary>
    public class SkinsModel {
        /// <summary>
        /// Skin Index
        /// </summary>
        public int SkinIndex { get; set; }
        /// <summary>
        /// Skin
        /// </summary>
        public SkinModel[] Skin { get; set; }
        /// <summary>
        /// Count
        /// </summary>
        public int Count { get; set; }
        /// <summary>
        /// Default Skin Index
        /// </summary>
        public int DefaultSkinIndex { get; set; }
    }
}