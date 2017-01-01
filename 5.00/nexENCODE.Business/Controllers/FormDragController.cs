using System.Drawing;
using System.Windows.Forms;
using nexENCODE.Models;
namespace nexENCODE.Business.Controllers {
    /// <summary>
    /// Form Drag Controller
    /// </summary>
    public class FormDragController {
        /// <summary>
        /// Model
        /// </summary>
        private FormDragModel _model;
        /// <summary>
        /// Form Drag Controller
        /// </summary>
        public FormDragController() {
            _model = new FormDragModel();
        }
        /// <summary>
        /// Form Mouse Down
        /// </summary>
        /// <param name="frm"></param>
        /// <param name="mousePos"></param>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void Form_MouseDown(Form frm, Point mousePos, object sender, System.Windows.Forms.MouseEventArgs e) {
            _model.DragPointA = mousePos.X - frm.Location.X;
            _model.DragPointB = mousePos.Y - frm.Location.Y;
        }
        /// <summary>
        /// Form Mouse Move
        /// </summary>
        /// <param name="frm"></param>
        /// <param name="mousePos"></param>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        public void Form_MouseMove(Form frm, Point mousePos, object sender, System.Windows.Forms.MouseEventArgs e) {
            if (e.Button == MouseButtons.Left) {
                _model.NewPoint = mousePos;
                _model.NewPoint.X = _model.NewPoint.X - _model.DragPointA;
                _model.NewPoint.Y = _model.NewPoint.Y - _model.DragPointB;
                frm.Location = _model.NewPoint;
            }
        }
    }
}