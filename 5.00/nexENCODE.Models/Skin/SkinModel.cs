using System.Collections.Generic;
namespace nexENCODE.Models.Skin {
    /// <summary>
    /// Skin Model
    /// </summary>
    public class SkinModel {
        /// <summary>
        /// File Name
        /// </summary>
        public string FileName { get; set; }
        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Icon
        /// </summary>
        public string Icon { get; set; }
        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }
        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }
        /// <summary>
        /// Combine
        /// </summary>
        public bool Combine { get; set; }
        /// <summary>
        /// Use Window Metrics
        /// </summary>
        public bool UseWindowMetrics { get; set; }
        /// <summary>
        /// Main Window Shape Count
        /// </summary>
        public int MainWindow_ShapeCount { get; set; }
        /// <summary>
        /// Main Window Parent Shape Region
        /// </summary>
        public int MainWindow_ParentShapeRegion { get; set; }
        /// <summary>
        /// Main Window Shape File Name
        /// </summary>
        public string MainWindow_ShapeFileName { get; set; }
        /// <summary>
        /// Main Window Background Image
        /// </summary>
        public string MainWindow_BackgroundImage { get; set; }
        /// <summary>
        /// Main Window Object Count
        /// </summary>
        public int MainWindow_ObjectCount { get; set; }
        /// <summary>
        /// Main Window Object File Name
        /// </summary>
        public string MainWindow_ObjectFileName { get; set; }
        /// <summary>
        /// Main Window Set Shape
        /// </summary>
        public bool MainWindow_SetShape { get; set; }
        /// <summary>
        /// Main Window Code File
        /// </summary>
        public string MainWindow_CodeFile { get; set; }
        /// <summary>
        /// Main Window Objects
        /// </summary>
        public ObjectModel[] MainWindow_Objects { get; set; }
        /// <summary>
        /// Main Window Shape
        /// </summary>
        public ShapeModel[] MainWindow_Shape { get; set; }
        /// <summary>
        /// Skin Model
        /// </summary>
        public SkinModel() {
        }
    }
}