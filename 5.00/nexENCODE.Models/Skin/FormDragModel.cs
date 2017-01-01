using System.Drawing;
namespace nexENCODE.Models.Skin {
    /// <summary>
    /// Form Drag Model
    /// </summary>
    public class FormDragModel {
        /// <summary>
        /// New Point
        /// </summary>
        public Point NewPoint;
        /// <summary>
        /// Drag Point A
        /// </summary>
        public int DragPointA { get; set; }
        /// <summary>
        /// Drag Point B
        /// </summary>
        public int DragPointB { get; set; }
    }
}