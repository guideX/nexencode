using nexENCODE.Enum.Skin;
using nexENCODE.Models.Skin;
using System.Windows.Forms;
using TeamNexgenCore.Helpers;
namespace nexENCODE.Business.Controllers {
    /// <summary>
    /// Skins Controller
    /// </summary>
    public class SkinsController {
        /// <summary>
        /// Skins
        /// </summary>
        private SkinsModel _skins { get; set; }
        /// <summary>
        /// Skins Controller
        /// </summary>
        public SkinsController() {
            _skins = new SkinsModel();
        }
        /// <summary>
        /// Return Main Window Code File
        /// </summary>
        /// <param name="skinIndex"></param>
        /// <returns></returns>
        public string MainWnd_CodeFile(int skinIndex) {
            return _skins.Skin[skinIndex].MainWindow_CodeFile;
        }
        /// <summary>
        /// Window Size
        /// </summary>
        /// <param name="type"></param>
        /// <param name="frm"></param>
        public void WindowSize(WindowSizes type, Form frm, string ini) {
            if (!string.IsNullOrEmpty(frm.Name)) {
                if (type == WindowSizes.Loading) {
                    frm.Left = IniFileHelper.ReadIniInt(ini, frm.Name, "Left", frm.Left);
                    frm.Top = IniFileHelper.ReadIniInt(ini, frm.Name, "Top", frm.Top);
                    frm.Width = IniFileHelper.ReadIniInt(ini, frm.Name, "Width", frm.Width);
                    frm.Height = IniFileHelper.ReadIniInt(ini, frm.Name, "Height", frm.Height);
                } else {
                    IniFileHelper.WriteIni(ini, frm.Name, "Left", frm.Left.ToString());
                    IniFileHelper.WriteIni(ini, frm.Name, "Top", frm.Top.ToString());
                    IniFileHelper.WriteIni(ini, frm.Name, "Width", frm.Width.ToString());
                    IniFileHelper.WriteIni(ini, frm.Name, "Height", frm.Height.ToString());
                }
            }
        }
        /// <summary>
        /// Read Last Skin Index
        /// </summary>
        /// <returns></returns>
        public int ReadIndex() {
            return _skins.SkinIndex;
        }
        /// <summary>
        /// Find
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public int? Find(string name) {
            for (var i = 0; i <= _skins.Skin.Length - 1; i++) {
                if (name.Trim().ToLower() == _skins.Skin[i].Name.Trim().ToLower()) {
                    return i;
                }
            }
            return null;
        }
        /// <summary>
        /// Replace Indicators
        /// </summary>
        /// <param name="path"></param>
        /// <param name="skinFile"></param>
        /// <returns></returns>
        public string ReplaceIndicators(string path, string skinFile = "") {
            path = path.Replace("$apppath", Application.StartupPath);
            path = path.Replace("$skinspath", Application.StartupPath + @"\data\skins");
            if (!string.IsNullOrWhiteSpace(skinFile)) {
                var msg = FilesController.ReadDirectory(path);
                path = path.Replace("$skinpath", msg);
            }
            return path;
        }
        /// <summary>
        /// Do Skin Files Exist
        /// </summary>
        /// <param name="skinIndex"></param>
        /// <returns></returns>
        public bool DoSkinFilesExist(int skinIndex) {
            if (skinIndex != 0) {
                if (!System.IO.File.Exists(_skins.Skin[skinIndex].FileName)) throw new System.Exception("'Skin File' not found");
                if (!System.IO.File.Exists(_skins.Skin[skinIndex].MainWindow_ShapeFileName)) throw new System.Exception("'Shape File' not found");
                if (!System.IO.File.Exists(_skins.Skin[skinIndex].MainWindow_ObjectFileName)) throw new System.Exception("'Object File' not found");
                return true;
            } else {
                return false;
            }
        }

        /*
        public bool LoadShape(Form lForm, int lSkinIndex) {
            clsAPI.gWindowSettings lWindowSettings = lAPI.WindowSettings(lForm);
            int X = 0;
            int Y = 0;
            int i = 0;
            clsAPI.eCombineRegionRet lCombineRegionRet = default(clsAPI.eCombineRegionRet);
            X = lWindowSettings.wWindowBorder;
            Y = lWindowSettings.wTitleBarHeight;
            var _with1 = lSkins.Skin(lSkinIndex);
            if (_with1.MainWindow_SetShape == false)
                return true;            if (_with1.MainWindow_ShapeCount != 0) {
                for (i = 1; i <= _with1.MainWindow_ShapeCount; i++) {
                    switch (_with1.MainWindow_Shape(i).Type) {
                        case ShapeTypes.RoundRectRgn:
                            if (_with1.UseWindowMetrics == true) {
                                _with1.MainWindow_Shape(i).Rgn.Rgn = lAPI.ReturnRegion(ShapeTypes.RoundRectRgn, X + _with1.MainWindow_Shape(i).Rgn.X1, Y + _with1.MainWindow_Shape(i).Rgn.Y1, X + _with1.MainWindow_Shape(i).Rgn.X2, Y + _with1.MainWindow_Shape(i).Rgn.Y2, _with1.MainWindow_Shape(i).Rgn.X3, _with1.MainWindow_Shape(i).Rgn.Y3);
                            } else {
                                _with1.MainWindow_Shape(i).Rgn.Rgn = lAPI.ReturnRegion(ShapeTypes.RoundRectRgn, _with1.MainWindow_Shape(i).Rgn.X1, _with1.MainWindow_Shape(i).Rgn.Y1, _with1.MainWindow_Shape(i).Rgn.X2, _with1.MainWindow_Shape(i).Rgn.Y2, _with1.MainWindow_Shape(i).Rgn.X3, _with1.MainWindow_Shape(i).Rgn.Y3);
                            }
                            break;
                        default:
                            if (_with1.UseWindowMetrics == true) {
                                _with1.MainWindow_Shape(i).Rgn.Rgn = lAPI.ReturnRegion(_with1.MainWindow_Shape(i).Type, X + _with1.MainWindow_Shape(i).Rgn.X1, Y + _with1.MainWindow_Shape(i).Rgn.Y1, X + _with1.MainWindow_Shape(i).Rgn.X2, Y + _with1.MainWindow_Shape(i).Rgn.Y2);
                            } else {
                                _with1.MainWindow_Shape(i).Rgn.Rgn = lAPI.ReturnRegion(_with1.MainWindow_Shape(i).Type, _with1.MainWindow_Shape(i).Rgn.X1, _with1.MainWindow_Shape(i).Rgn.Y1, _with1.MainWindow_Shape(i).Rgn.X2, _with1.MainWindow_Shape(i).Rgn.Y2);
                            }
                            break;
                    }
                }
                if (_with1.Combine == true) {
                    for (i = 1; i <= _with1.MainWindow_ShapeCount; i++) {
                        if (_with1.MainWindow_Shape(i).CombineMode != 0 & _with1.MainWindow_Shape(i).DestRgn != 0 & _with1.MainWindow_Shape(i).SrcRgn1 != 0 & _with1.MainWindow_Shape(i).SrcRgn2 != 0) {
                            lCombineRegionRet = lAPI.CombineRegion(_with1.MainWindow_Shape(_with1.MainWindow_Shape(i).DestRgn).Rgn.Rgn, _with1.MainWindow_Shape(_with1.MainWindow_Shape(i).SrcRgn1).Rgn.Rgn, _with1.MainWindow_Shape(_with1.MainWindow_Shape(i).SrcRgn2).Rgn.Rgn, _with1.MainWindow_Shape(i).CombineMode);
                            if (lCombineRegionRet != clsAPI.eCombineRegionRet.cSimpleRegion & lCombineRegionRet != clsAPI.eCombineRegionRet.cComplexRegion & lCombineRegionRet != clsAPI.eCombineRegionRet.cNullRegion) {
                                if (ProcessError != null) {
                                    ProcessError(lAPI.lLastError, "CombineRegion");
                                }
                            }
                        }
                    }
                }
                return lAPI.SetWindowRegion(lForm, _with1.MainWindow_Shape(_with1.MainWindow_ParentShapeRegion).Rgn.Rgn, true);
            } else {
                return false;
            }
        }*/

        /// <summary>
        /// Load Objects
        /// </summary>
        /// <returns></returns>
        //public bool LoadObjects(Form frm, int skinIndex) {

        //}
    }
}