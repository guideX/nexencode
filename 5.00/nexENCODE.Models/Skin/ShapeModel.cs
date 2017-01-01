using nexENCODE.Enum.Skin;
namespace nexENCODE.Models.Skin {
    /// <summary>
    /// Shape Model
    /// </summary>
    public class ShapeModel {
        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Type
        /// </summary>
        public ShapeTypes Type { get; set; }
        /// <summary>
        /// Rgn
        /// </summary>
        public RegionModel Rgn { get; set; }
        /// <summary>
        /// Combine Mode
        /// </summary>
        public CombineModes CombineMode { get; set; }
        /// <summary>
        /// Dest Rgn
        /// </summary>
        public int DestRgn { get; set; }
        /// <summary>
        /// Src Rgn 1
        /// </summary>
        public int SrcRgn1 { get; set; }
        /// <summary>
        /// Src Rgn 2
        /// </summary>
        public int SrcRgn2 { get; set; }
    }
}