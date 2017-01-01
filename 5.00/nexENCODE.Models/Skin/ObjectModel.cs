using nexENCODE.Enum;
namespace nexENCODE.Models.Skin {
    /// <summary>
    /// Object Model
    /// </summary>
    public class ObjectModel {
        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Button Type
        /// </summary>
        public ButtonTypes ButtonType { get; set; }
        /// <summary>
        /// Label Type
        /// </summary>
        public LabelTypes LabelType { get; set; }
        /// <summary>
        /// Object Type
        /// </summary>
        public ObjectTypes ObjectType { get; set; }
        /// <summary>
        /// File Name
        /// </summary>
        public string Filename { get; set; }
        /// <summary>
        /// File Name 2
        /// </summary>
        public string Filename2 { get; set; }
        /// <summary>
        /// File Name 3
        /// </summary>
        public string Filename3 { get; set; }
        /// <summary>
        /// Left
        /// </summary>
        public int Left { get; set; }
        /// <summary>
        /// Top
        /// </summary>
        public int Top { get; set; }
        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }
        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }
        /// <summary>
        /// Visible
        /// </summary>
        public bool Visible { get; set; }
        /// <summary>
        /// On Click
        /// </summary>
        public string OnClick { get; set; }
    }
}