using DevExpress.Data;
using DevExpress.Utils;
using DevExpress.XtraCharts;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using JPlatform.Client.Controls6;
using JPlatform.Client.CSIGMESBaseform6;
using JPlatform.Client.JBaseForm6;
using JPlatform.Client.Library6.interFace;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;


namespace CSI.GMES.PD
{
    public partial class MSPD90226A : CSIGMESBaseform6//JERPBaseForm
    {
        public bool _firstLoad = true, _isMouseDown = false;
        public MyCellMergeHelper _Helper = null;
        public DataTable _dtSelected = null, _dtTotal = null;
        public GridHitInfo hitInfoStart = null;
        public int _tab = 0;

        public MSPD90226A()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            _firstLoad = true;

            base.OnLoad(e);
            NewButton = false;
            AddButton = false;
            DeleteRowButton = false;
            SaveButton = true;
            DeleteButton = true;
            PreviewButton = false;
            PrintButton = false;

            cboWorkDate.EditValue = DateTime.Now.ToString();
            cboAssDate.EditValue = DateTime.Now.ToString();
            gvwMain.OptionsSelection.MultiSelect = true;
            panTop.BackColor = Color.FromArgb(240, 240, 240);
            this.lblSave.Font = new System.Drawing.Font("Times New Roman", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSave.ForeColor = System.Drawing.Color.White;
            InitCombobox();

            _firstLoad = false;
        }

        #region [Start Button Event Code By UIBuilder]

        public override void QueryClick()
        {
            try
            {
                pbProgressShow();

                if (_tab == 0)
                {
                    CreateTableSelect();
                    InitControls(grdMain);
                    DataTable _dtSource = GetData("Q");
                    DataTable _dtMap = GetData("Q_TOTAL");

                    if (_dtMap != null && _dtMap.Rows.Count > 0)
                    {
                        _dtTotal = _dtMap.Copy();
                    }
                    else
                    {
                        _dtTotal = null;
                    }

                    if (_dtSource != null && _dtSource.Rows.Count > 0)
                    {
                        var _distinctValues = _dtSource.AsEnumerable()
                                            .Select(row => new
                                            {
                                                PFC_PART_NO = row.Field<string>("PFC_PART_NO"),
                                                COMPONENT_NM = row.Field<string>("COMPONENT_NM"),
                                                COMPONENT_NM_VN = row.Field<string>("COMPONENT_NM_VN"),
                                                ORD = row.Field<decimal>("ORD"),
                                            })
                                            .Distinct().OrderBy(r => r.ORD);
                        DataTable _dtHead = LINQResultToDataTable(_distinctValues).Select("", "ORD").CopyToDataTable();
                        CreateSizeGrid(grdMain, gvwMain, _dtSource, _dtHead);
                        DataTable _dtf = Binding_Data(_dtSource, gvwMain);
                        SetData(grdMain, _dtf);
                        Formart_Grid_Main(grdMain, gvwMain);
                    }
                    else
                    {
                        grdMain.DataSource = null;
                        gvwMain.Columns.Clear();
                        gvwMain.Bands.Clear();
                    }
                }
                else if(_tab == 1)
                {
                    DataTable _dtSource = GetData("Q_CONFIRM");

                    if(_dtSource != null && _dtSource.Rows.Count > 0)
                    {
                        DataTable _dtContent = _dtSource.Select("CS_SIZE <> 'G-TOTAL'", "MLINE_CD, ASY_YMD, STYLE_CD, GRP_NO").CopyToDataTable();
                        DataTable _dtHead = _dtSource.Select("CS_SIZE = 'G-TOTAL'", "ORD").CopyToDataTable();
                        CreateSizeGrid(grdConfirm, gvwConfirm, _dtContent, _dtHead);
                        DataTable _dtf = Binding_Data(_dtContent, gvwConfirm);
                        SetData(grdConfirm, _dtf);
                        Formart_Grid_Main(grdConfirm, gvwConfirm);
                    }
                    else
                    {
                        grdConfirm.DataSource = null;
                        gvwConfirm.Columns.Clear();
                        gvwConfirm.Bands.Clear();
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                pbProgressHide();
            }
        }

        public void CreateTableSelect()
        {
            _dtSelected = null;
            _dtSelected = new DataTable();

            _dtSelected.Columns.Add("LINE_CD", typeof(String));
            _dtSelected.Columns.Add("MLINE_CD", typeof(String));
            _dtSelected.Columns.Add("INPUT_PRIO", typeof(String));
            _dtSelected.Columns.Add("STYLE_CD", typeof(String));
            _dtSelected.Columns.Add("CS_SIZE", typeof(String));
            _dtSelected.Columns.Add("PFC_PART_NO", typeof(String));
            _dtSelected.Columns.Add("DIR_QTY", typeof(String));
        }
        public override void DeleteClick()
        {
            base.DeleteClick();
            try
            {
                DialogResult dlr;

                //string _machine = string.IsNullOrEmpty(cboMachine.EditValue.ToString()) ? "" : cboMachine.EditValue.ToString();

                string assy_ymd_tmp = cboAssDate.yyyymmdd;
                string assy_ymd = assy_ymd_tmp.Substring(0, 4) + "-" + assy_ymd_tmp.Substring(4, 2) + "-" + assy_ymd_tmp.Substring(6, 2);

                dlr = MessageBox.Show("Bạn có muốn delete kế hoạch ngày assembly:" + assy_ymd + " không?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlr == DialogResult.Yes)
                {
                    bool result = SaveData("Q_DELETE");
                    if (result)
                    {
                        MessageBoxW("Delete successfully!", IconType.Information);
                        QueryClick();
                    }
                    else
                    {
                        MessageBoxW("Delete failed!", IconType.Warning);
                    }
                }

            }
            catch (Exception)
            {

                throw;
            }
          

        }
       
       
        public override void SaveClick()
        {
            try
            {
                DialogResult dlr;

                string _style = string.IsNullOrEmpty(cboStyle.EditValue.ToString()) ? "" : cboStyle.EditValue.ToString();
                string _machine = string.IsNullOrEmpty(cboMachine.EditValue.ToString()) ? "" : cboMachine.EditValue.ToString();
                string _hms = string.IsNullOrEmpty(cboHMS.EditValue.ToString()) ? "" : cboHMS.EditValue.ToString();
                string _part = chkcboPart.EditValue == null ? "" : chkcboPart.EditValue.ToString().Replace(" ", "");

                if (string.IsNullOrEmpty(_style))
                {
                    MessageBox.Show("Mã hàng không được trống!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //if (string.IsNullOrEmpty(_machine))
                //{
                //    MessageBox.Show("Mã máy không được trống!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    return;
                //}

                if (string.IsNullOrEmpty(_part))
                {
                    MessageBox.Show("Mã chi tiết không được trống!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if(_dtSelected != null && _dtSelected.Rows.Count < 1)
                {
                    MessageBox.Show("Chi tiết được chọn không được trống!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                dlr = MessageBox.Show("Bạn có muốn Save không?", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlr == DialogResult.Yes)
                {
                    bool result = SaveData("Q_SAVE");
                    if (result)
                    {
                        MessageBoxW("Save successfully!", IconType.Information);
                        QueryClick();
                    }
                    else
                    {
                        MessageBoxW("Save failed!", IconType.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public DataTable FormatDataTable(DataTable _dtSource)
        {
            DataTable _dtRef = _dtSource;

            for (int iRow = 0; iRow < _dtRef.Rows.Count; iRow++)
            {
                _dtRef.Rows[iRow]["QTY"] = Math.Round(Double.Parse(_dtRef.Rows[iRow]["QTY"].ToString()),1);
            }

            return _dtRef;
        }

        public DataTable Binding_Data(DataTable dtSource, BandedGridViewEx gridView)
        {
            try
            {
                DataTable _dtf = GetDataTable(gridView);
                string _col_nm = "", _distinct_row = "";

                for (int iRow = 0; iRow < dtSource.Rows.Count; iRow++)
                {
                    if (!dtSource.Rows[iRow]["DISTINCT_ROW"].ToString().Equals(_distinct_row))
                    {
                        _dtf.Rows.Add();

                        _dtf.Rows[_dtf.Rows.Count - 1]["LINE_CD"] = dtSource.Rows[iRow]["LINE_CD"].ToString();
                        _dtf.Rows[_dtf.Rows.Count - 1]["LINE_NM"] = dtSource.Rows[iRow]["LINE_NM"].ToString();
                        _dtf.Rows[_dtf.Rows.Count - 1]["MLINE_CD"] = dtSource.Rows[iRow]["MLINE_CD"].ToString();
                        _dtf.Rows[_dtf.Rows.Count - 1]["MODEL_NM"] = dtSource.Rows[iRow]["MODEL_NM"].ToString();
                        _dtf.Rows[_dtf.Rows.Count - 1]["STYLE_CD"] = dtSource.Rows[iRow]["STYLE_CD"].ToString();
                        _dtf.Rows[_dtf.Rows.Count - 1]["CS_SIZE"] = dtSource.Rows[iRow]["CS_SIZE"].ToString();

                        if (_tab == 0)
                        {
                            _dtf.Rows[_dtf.Rows.Count - 1]["INPUT_PRIO"] = dtSource.Rows[iRow]["INPUT_PRIO"].ToString();
                        }
                        else if(_tab == 1)
                        {
                            _dtf.Rows[_dtf.Rows.Count - 1]["ASY_YMD"] = dtSource.Rows[iRow]["ASY_YMD"].ToString();
                        }

                        _distinct_row = dtSource.Rows[iRow]["DISTINCT_ROW"].ToString();
                    }

                    _col_nm = dtSource.Rows[iRow]["PFC_PART_NO"].ToString();

                    if (_tab == 0)
                    {
                        if (dtSource.Rows[iRow]["REGISTER_YN"].ToString().Equals("Y"))
                        {
                            _dtf.Rows[_dtf.Rows.Count - 1][_col_nm] = dtSource.Rows[iRow]["DIR_QTY"].ToString() + "_SAVED";
                        }
                        else
                        {
                            _dtf.Rows[_dtf.Rows.Count - 1][_col_nm] = dtSource.Rows[iRow]["DIR_QTY"].ToString();
                        }
                    }
                    else if (_tab == 1)
                    {
                        _dtf.Rows[_dtf.Rows.Count - 1][_col_nm] = dtSource.Rows[iRow]["DIR_QTY"].ToString();
                    }
                }

                return _dtf;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public void CreateSizeGrid(GridControlEx gridControl, BandedGridViewEx gridView, DataTable dtSource, DataTable dtHead)
        {
            gridView.BeginDataUpdate();
            try
            {
                gridControl.DataSource = null;
                InitControls(gridControl);
                gridView.Columns.Clear();
                gridView.Bands.Clear();

                while (gridView.Columns.Count > 0)
                {
                    gridView.Columns.RemoveAt(0);
                }

                gridView.OptionsView.ShowColumnHeaders = false; 

                GridBandEx gridBand = null, gridBandChild = null;
                BandedGridColumnEx colBand = new BandedGridColumnEx();
                int _col_start = Int32.Parse(dtSource.Rows[0]["COL_START"].ToString());
                int _col_end = Int32.Parse(dtSource.Rows[0]["COL_END"].ToString());
                string _caption = _tab == 0 ? "Total Saved" : "G-Total";

                for (int iCol = 0; iCol <= _col_end; iCol++)
                {
                    ////////// Column
                    gridBand = new GridBandEx() { Caption = Get_Column_Caption(dtSource.Columns[iCol].ColumnName.ToString()) };
                    gridView.Bands.Add(gridBand);
                    gridBand.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gridBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap;
                    gridBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                    gridBand.AppearanceHeader.Options.UseBackColor = true;
                    gridBand.RowCount = 2;
                    gridBand.Visible = iCol < _col_start ? false : true;

                    gridBandChild = new GridBandEx() { Caption = dtSource.Columns[iCol].ColumnName.ToString().Equals("LINE_NM") ? _caption : "" };
                    gridBandChild.AppearanceHeader.Options.UseTextOptions = true;
                    gridBandChild.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridBandChild.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridBandChild.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                    gridBandChild.Name = "";
                    gridBandChild.VisibleIndex = 0;

                    gridBand.Children.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] { gridBandChild  });

                    colBand = new BandedGridColumnEx() { FieldName = dtSource.Columns[iCol].ColumnName.ToString(), Visible = true };
                    colBand.Width = 100;
                    gridBandChild.Columns.Add(colBand);
                }

                //////// Size Column
                for (int iRow = 0; iRow < dtHead.Rows.Count; iRow++)
                {
                    gridBand = new GridBandEx() { Caption = dtHead.Rows[iRow]["COMPONENT_NM"].ToString() + "\n(" + dtHead.Rows[iRow]["COMPONENT_NM_VN"].ToString() + ")" };
                    gridView.Bands.Add(gridBand);
                    gridBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap;
                    gridBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                    gridBand.AppearanceHeader.Options.UseBackColor = true;
                    gridBand.RowCount = 2;
                    gridBand.Visible = true;

                    if (_tab == 0)
                    {
                        gridBandChild = new GridBandEx() { Caption = "" };
                    }
                    else if (_tab == 1)
                    {
                        gridBandChild = new GridBandEx() { Caption = getDataTotal(dtHead.Rows[iRow]["PFC_PART_NO"].ToString(), dtHead) };
                    }

                    gridBandChild.AppearanceHeader.Options.UseTextOptions = true;
                    gridBandChild.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                    gridBandChild.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridBandChild.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                    gridBandChild.Name = "";
                    gridBandChild.VisibleIndex = 0;

                    gridBand.Children.AddRange(new DevExpress.XtraGrid.Views.BandedGrid.GridBand[] { gridBandChild });

                    colBand = new BandedGridColumnEx() { FieldName = dtHead.Rows[iRow]["PFC_PART_NO"].ToString(), Visible = true };
                    colBand.Width = 100;
                    gridBandChild.Columns.Add(colBand);
                }
            }
            catch
            {
                //throw EX;
            }
            gridView.EndDataUpdate();
            gridView.ExpandAllGroups();
        }

        public string Get_Column_Caption(string _type)
        {
            string _result = "";

            switch (_type)
            {
                case "LINE_NM":
                    _result = "Line";
                    break;
                case "MLINE_CD":
                    _result = "Mini Line";
                    break;
                case "INPUT_PRIO":
                    _result = "Hour";
                    break;
                case "MODEL_NM":
                    _result = "Model Name";
                    break;
                case "STYLE_CD":
                    _result = "Style Code";
                    break;
                case "CS_SIZE":
                    _result = "Size";
                    break;
                case "ASY_YMD":
                    _result = "Assembly Date";
                    break;
                default:
                    break;
            }

            return _result;
        }


        public void Formart_Grid_Main(GridControlEx gridControl, BandedGridViewEx gridVieW)
        {
            try
            {
                gridControl.BeginUpdate();

                for (int i = 0; i < gridVieW.Columns.Count; i++)
                {
                    gridVieW.Columns[i].OptionsColumn.AllowEdit = false;
                    gridVieW.Columns[i].OptionsColumn.ReadOnly = true;
                    gridVieW.Columns[i].OptionsColumn.AllowSort = DefaultBoolean.False;
                    gridVieW.Columns[i].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridVieW.Columns[i].AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridVieW.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gridVieW.Columns[i].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gridVieW.Columns[i].AppearanceCell.TextOptions.WordWrap = WordWrap.Wrap;
                    gridVieW.Columns[i].AppearanceCell.Font = new System.Drawing.Font("Calibri", 12, FontStyle.Regular);

                    gridVieW.OptionsSelection.MultiSelect = true;
                    gridVieW.OptionsSelection.MultiSelectMode = GridMultiSelectMode.RowSelect;
                    gridVieW.OptionsSelection.EnableAppearanceHideSelection = true;
                    gridVieW.OptionsSelection.EnableAppearanceFocusedCell = true;

                    if (i >= 7)
                    {
                        gridVieW.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                        gridVieW.Columns[i].DisplayFormat.FormatType = FormatType.Numeric;
                        gridVieW.Columns[i].DisplayFormat.FormatString = "#,#0.#";
                    }

                    string _col_nm = gridVieW.Columns[i].FieldName.ToString();
                    switch (_col_nm)
                    {
                        case "MLINE_CD":
                        case "INPUT_PRIO":
                        case "CS_SIZE":
                            gridVieW.Columns[i].Width = 60;
                            break;
                         case "MODEL_NM":
                            gridVieW.Columns[i].Width = 150;
                            gridVieW.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                            gridVieW.Columns[i].ColumnEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
                            break;
                        default:
                            break;
                    }
                }

                for (int j = 0; j < gridVieW.Columns.Count; j++)
                {
                    if (gridVieW.Columns[j].OwnerBand != null)
                    {
                        gridVieW.Columns[j].OwnerBand.AppearanceHeader.BackColor = Color.FromArgb(255, 250, 179);

                        if (j >= 7)
                        {
                            if(_tab == 0)
                            {
                                gridVieW.Columns[j].OwnerBand.Caption = getDataTotal(gridVieW.Columns[j].FieldName.ToString());
                            }
                        }
                    }
                }

                gridVieW.RowHeight = 30;
                gridControl.EndUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public string getDataTotal(string _col_nm, DataTable _dtf = null)
        {
            string _result = "0";

            if (_tab == 0)
            {
                if (_dtTotal == null || _dtTotal.Rows.Count < 1) return "0";

                for (int iRow = 0; iRow < _dtTotal.Rows.Count; iRow++)
                {
                    if (_dtTotal.Rows[iRow]["PFC_PART_NO"].ToString().Equals(_col_nm))
                    {
                        _result = _dtTotal.Rows[iRow]["QTY"].ToString();
                        break;
                    }
                }
            }
            else if (_tab == 1)
            {
                if (_dtf == null || _dtf.Rows.Count < 1) return "0";

                for (int iRow = 0; iRow < _dtf.Rows.Count; iRow++)
                {
                    if (_dtf.Rows[iRow]["PFC_PART_NO"].ToString().Equals(_col_nm))
                    {
                        _result = _dtf.Rows[iRow]["DIR_QTY"].ToString();
                        break;
                    }
                }
            }

            return _result;
        }

        public string FormatNumber(string value)
        {
            return Math.Round(Double.Parse(value), 1).ToString();
        }

        #endregion [Start Button Event Code By UIBuilder] 

        #region [Grid]

        private DataTable GetData(string argType)
        {
            try
            {
                P_MSPD90226A_Q proc = new P_MSPD90226A_Q();
                DataTable dtData = null;

                string _factory = string.IsNullOrEmpty(cboFactory.EditValue.ToString()) ? "" : cboFactory.EditValue.ToString();
                string _line = string.IsNullOrEmpty(cboPlant.EditValue.ToString()) ? "" : cboPlant.EditValue.ToString();
                string _mline = string.IsNullOrEmpty(cboLine.EditValue.ToString()) ? "" : cboLine.EditValue.ToString();
                string _style = string.IsNullOrEmpty(cboStyle.EditValue.ToString()) ? "" : cboStyle.EditValue.ToString();
                string _machine = string.IsNullOrEmpty(cboMachine.EditValue.ToString()) ? "" : cboMachine.EditValue.ToString();
                string _hms = string.IsNullOrEmpty(cboHMS.EditValue.ToString()) ? "" : cboHMS.EditValue.ToString();
                string _part = chkcboPart.EditValue == null ? "" : chkcboPart.EditValue.ToString().Replace(" ", "");

                dtData = proc.SetParamData(dtData, argType, _factory, _line, _mline, _style, cboWorkDate.yyyymmdd, cboAssDate.yyyymmdd, _machine, _hms, _part);
                ResultSet rs = CommonCallQuery(dtData, proc.ProcName, proc.GetParamInfo(), false, 90000, "", true);
                if (rs == null || rs.ResultDataSet == null || rs.ResultDataSet.Tables.Count == 0 || rs.ResultDataSet.Tables[0].Rows.Count == 0)
                {
                    return null;
                }
                return rs.ResultDataSet.Tables[0];
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }


        #endregion [Grid]

        #region [Combobox]

        private void InitCombobox()
        {
            if (_tab == 0)
            {
                LoadDataCbo(cboFactory, "Factory", "Q_FTY");
                LoadDataCbo(cboPlant, "Plant", "Q_PLANT");
                LoadDataCbo(cboLine, "Line", "Q_LINE");

                LoadDataCbo(cboHMS, "Hour", "Q_HMS");
                LoadDataCbo(cboMachine, "Machine", "Q_MACHINE");
                LoadDataCbo(cboStyle, "Style Name", "Q_STYLE");
                LoadDataCbo(null, "Part", "Q_PART");
            }
            else if (_tab == 1)
            {
                LoadDataCbo(cboPlant, "Plant", "Q_PLANT_ALL");
                LoadDataCbo(cboLine, "Line", "Q_LINE_ALL");

                LoadDataCbo(cboHMS, "Hour", "Q_HMS_ALL");
                LoadDataCbo(cboMachine, "Machine", "Q_MACHINE");
                LoadDataCbo(cboStyle, "Style Name", "Q_STYLE_ALL");
            }
        }

        private void LoadDataCbo(LookUpEditEx argCbo, string _cbo_nm, string _type, string _search = "")
        {
            try
            {
                DataTable dt = Get_Data_Combobox(_type, _search);

                if (_type.Equals("Q_PART"))
                {
                    chkcboPart.Properties.Items.Clear();
                    chkcboPart.Properties.DataSource = null;
                    if (dt == null) return;

                    chkcboPart.Properties.SeparatorChar = '|';
                    for (int iRow = 0; iRow < dt.Rows.Count; iRow++)
                    {
                        chkcboPart.Properties.Items.Add(dt.Rows[iRow]["CODE"].ToString(), dt.Rows[iRow]["NAME"].ToString(), System.Windows.Forms.CheckState.Checked, true);
                    }
                }
                else
                {

                    if (dt == null || dt.Rows.Count < 1)
                    {
                        argCbo.Properties.Columns.Clear();
                        argCbo.Properties.DataSource = null;

                        return;
                    }

                    string columnCode = dt.Columns[0].ColumnName;
                    string columnName = dt.Columns[1].ColumnName;
                    string captionCode = "Code";
                    string captionName = _cbo_nm;

                    argCbo.Properties.Columns.Clear();
                    argCbo.Properties.DataSource = dt;
                    argCbo.Properties.ValueMember = columnCode;
                    argCbo.Properties.DisplayMember = columnName;
                    argCbo.Properties.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo(columnCode));
                    argCbo.Properties.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo(columnName));
                    argCbo.Properties.Columns[columnCode].Visible = _type.Contains("Q_STYLE") ? true : false;
                    argCbo.Properties.Columns[columnCode].Width = 10;
                    argCbo.Properties.Columns[columnCode].Caption = captionCode;
                    argCbo.Properties.Columns[columnName].Caption = captionName;
                    argCbo.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        public DataTable Get_Data_Combobox(string _type, string _search)
        {
            try
            {
                P_MSPD90226A_COMBO proc = new P_MSPD90226A_COMBO();
                DataTable dtData = null;

                string _factory = string.IsNullOrEmpty(cboFactory.EditValue.ToString()) ? "" : cboFactory.EditValue.ToString();
                string _line = string.IsNullOrEmpty(cboPlant.EditValue.ToString()) ? "" : cboPlant.EditValue.ToString();
                string _mline = string.IsNullOrEmpty(cboLine.EditValue.ToString()) ? "" : cboLine.EditValue.ToString();
                string _style = string.IsNullOrEmpty(cboStyle.EditValue.ToString()) ? "" : cboStyle.EditValue.ToString();
                string _hms = string.IsNullOrEmpty(cboHMS.EditValue.ToString()) ? "" : cboHMS.EditValue.ToString();
                string _machine = string.IsNullOrEmpty(cboMachine.EditValue.ToString()) ? "" : cboMachine.EditValue.ToString();

                dtData = proc.SetParamData(dtData, _type, _factory, _line, _mline, _style, cboWorkDate.yyyymmdd, cboAssDate.yyyymmdd, _hms, _machine);
                ResultSet rs = CommonCallQuery(dtData, proc.ProcName, proc.GetParamInfo(), false, 90000, "", true);

                if (rs == null || rs.ResultDataSet == null || rs.ResultDataSet.Tables.Count == 0 || rs.ResultDataSet.Tables[0].Rows.Count == 0)
                {
                    return null;
                }
                return rs.ResultDataSet.Tables[0];
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        #endregion [Combobox]

        #region Events

        private void gvwMain_CellMerge(object sender, CellMergeEventArgs e)
        {
            try
            {
                if (grdMain.DataSource == null || gvwMain.RowCount < 1) return;

                e.Merge = false;
                e.Handled = true;

                if (e.Column.FieldName.ToString().Equals("LINE_NM"))
                {
                    string _value1 = gvwMain.GetRowCellValue(e.RowHandle1, e.Column.FieldName.ToString()).ToString();
                    string _value2 = gvwMain.GetRowCellValue(e.RowHandle2, e.Column.FieldName.ToString()).ToString();

                    if (_value1 == _value2 && !string.IsNullOrEmpty(_value1))
                    {
                        e.Merge = true;
                    }
                }

                if (e.Column.FieldName.ToString().Equals("MLINE_CD"))
                {
                    string _value1 = gvwMain.GetRowCellValue(e.RowHandle1, "LINE_NM").ToString();
                    string _value2 = gvwMain.GetRowCellValue(e.RowHandle2, "LINE_NM").ToString();
                    string _value3 = gvwMain.GetRowCellValue(e.RowHandle1, e.Column.FieldName.ToString()).ToString();
                    string _value4 = gvwMain.GetRowCellValue(e.RowHandle2, e.Column.FieldName.ToString()).ToString();

                    if (_value1 == _value2 && !string.IsNullOrEmpty(_value1) &&
                        _value3 == _value4 && !string.IsNullOrEmpty(_value3))
                    {
                        e.Merge = true;
                    }
                }

                if (e.Column.FieldName.ToString().Equals("INPUT_PRIO"))
                {
                    string _value1 = gvwMain.GetRowCellValue(e.RowHandle1, "LINE_NM").ToString();
                    string _value2 = gvwMain.GetRowCellValue(e.RowHandle2, "LINE_NM").ToString();
                    string _value3 = gvwMain.GetRowCellValue(e.RowHandle1, "MLINE_CD").ToString();
                    string _value4 = gvwMain.GetRowCellValue(e.RowHandle2, "MLINE_CD").ToString();
                    string _value5 = gvwMain.GetRowCellValue(e.RowHandle1, e.Column.FieldName.ToString()).ToString();
                    string _value6 = gvwMain.GetRowCellValue(e.RowHandle2, e.Column.FieldName.ToString()).ToString();

                    if (_value1 == _value2 && !string.IsNullOrEmpty(_value1) &&
                        _value3 == _value4 && !string.IsNullOrEmpty(_value3) &&
                        _value5 == _value6 && !string.IsNullOrEmpty(_value5))
                    {
                        e.Merge = true;
                    }
                }

                if (e.Column.FieldName.ToString().Equals("MODEL_NM") || e.Column.FieldName.ToString().Equals("STYLE_CD"))
                {
                    string _value1 = gvwMain.GetRowCellValue(e.RowHandle1, "LINE_NM").ToString();
                    string _value2 = gvwMain.GetRowCellValue(e.RowHandle2, "LINE_NM").ToString();
                    string _value3 = gvwMain.GetRowCellValue(e.RowHandle1, "MLINE_CD").ToString();
                    string _value4 = gvwMain.GetRowCellValue(e.RowHandle2, "MLINE_CD").ToString();
                    string _value5 = gvwMain.GetRowCellValue(e.RowHandle1, "INPUT_PRIO").ToString();
                    string _value6 = gvwMain.GetRowCellValue(e.RowHandle2, "INPUT_PRIO").ToString();
                    string _value7 = gvwMain.GetRowCellValue(e.RowHandle1, e.Column.FieldName.ToString()).ToString();
                    string _value8 = gvwMain.GetRowCellValue(e.RowHandle2, e.Column.FieldName.ToString()).ToString();

                    if (_value1 == _value2 && !string.IsNullOrEmpty(_value1) &&
                        _value3 == _value4 && !string.IsNullOrEmpty(_value3) &&
                        _value5 == _value6 && !string.IsNullOrEmpty(_value5) &&
                        _value7 == _value8 && !string.IsNullOrEmpty(_value7))
                    {
                        e.Merge = true;
                    }
                }
            }
            catch { }
        }

        private void gvwConfirm_CellMerge(object sender, CellMergeEventArgs e)
        {
            try
            {
                if (grdConfirm.DataSource == null || gvwConfirm.RowCount < 1) return;

                e.Merge = false;
                e.Handled = true;

                if (e.Column.FieldName.ToString().Equals("LINE_NM"))
                {
                    string _value1 = gvwConfirm.GetRowCellValue(e.RowHandle1, e.Column.FieldName.ToString()).ToString();
                    string _value2 = gvwConfirm.GetRowCellValue(e.RowHandle2, e.Column.FieldName.ToString()).ToString();

                    if (_value1 == _value2 && !string.IsNullOrEmpty(_value1))
                    {
                        e.Merge = true;
                    }
                }

                if (e.Column.FieldName.ToString().Equals("MLINE_CD"))
                {
                    string _value1 = gvwConfirm.GetRowCellValue(e.RowHandle1, "LINE_NM").ToString();
                    string _value2 = gvwConfirm.GetRowCellValue(e.RowHandle2, "LINE_NM").ToString();
                    string _value3 = gvwConfirm.GetRowCellValue(e.RowHandle1, e.Column.FieldName.ToString()).ToString();
                    string _value4 = gvwConfirm.GetRowCellValue(e.RowHandle2, e.Column.FieldName.ToString()).ToString();

                    if (_value1 == _value2 && !string.IsNullOrEmpty(_value1) &&
                        _value3 == _value4 && !string.IsNullOrEmpty(_value3))
                    {
                        e.Merge = true;
                    }
                }

                if (e.Column.FieldName.ToString().Equals("ASY_YMD"))
                {
                    string _value1 = gvwConfirm.GetRowCellValue(e.RowHandle1, "LINE_NM").ToString();
                    string _value2 = gvwConfirm.GetRowCellValue(e.RowHandle2, "LINE_NM").ToString();
                    string _value3 = gvwConfirm.GetRowCellValue(e.RowHandle1, "MLINE_CD").ToString();
                    string _value4 = gvwConfirm.GetRowCellValue(e.RowHandle2, "MLINE_CD").ToString();
                    string _value5 = gvwConfirm.GetRowCellValue(e.RowHandle1, e.Column.FieldName.ToString()).ToString();
                    string _value6 = gvwConfirm.GetRowCellValue(e.RowHandle2, e.Column.FieldName.ToString()).ToString();

                    if (_value1 == _value2 && !string.IsNullOrEmpty(_value1) &&
                        _value3 == _value4 && !string.IsNullOrEmpty(_value3) &&
                        _value5 == _value6 && !string.IsNullOrEmpty(_value5))
                    {
                        e.Merge = true;
                    }
                }

                if (e.Column.FieldName.ToString().Equals("MODEL_NM") || e.Column.FieldName.ToString().Equals("STYLE_CD"))
                {
                    string _value1 = gvwConfirm.GetRowCellValue(e.RowHandle1, "LINE_NM").ToString();
                    string _value2 = gvwConfirm.GetRowCellValue(e.RowHandle2, "LINE_NM").ToString();
                    string _value3 = gvwConfirm.GetRowCellValue(e.RowHandle1, "MLINE_CD").ToString();
                    string _value4 = gvwConfirm.GetRowCellValue(e.RowHandle2, "MLINE_CD").ToString();
                    string _value5 = gvwConfirm.GetRowCellValue(e.RowHandle1, "ASY_YMD").ToString();
                    string _value6 = gvwConfirm.GetRowCellValue(e.RowHandle2, "ASY_YMD").ToString();
                    string _value7 = gvwConfirm.GetRowCellValue(e.RowHandle1, e.Column.FieldName.ToString()).ToString();
                    string _value8 = gvwConfirm.GetRowCellValue(e.RowHandle2, e.Column.FieldName.ToString()).ToString();

                    if (_value1 == _value2 && !string.IsNullOrEmpty(_value1) &&
                        _value3 == _value4 && !string.IsNullOrEmpty(_value3) &&
                        _value5 == _value6 && !string.IsNullOrEmpty(_value5) &&
                        _value7 == _value8 && !string.IsNullOrEmpty(_value7))
                    {
                        e.Merge = true;
                    }
                }
            }
            catch { }
        }

        private void cboFactory_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                if (_tab == 0)
                {
                    LoadDataCbo(cboPlant, "Plant", "Q_PLANT");
                }
                else if(_tab == 1)
                {
                    LoadDataCbo(cboPlant, "Plant", "Q_PLANT_ALL");
                }
            }
        }

        private void cboPlant_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                if (_tab == 0)
                {
                    LoadDataCbo(cboLine, "Line", "Q_LINE");
                    LoadDataCbo(cboStyle, "Style Name", "Q_STYLE");
                    LoadDataCbo(null, "Part", "Q_PART");
                }
                else if(_tab == 1)
                {
                    LoadDataCbo(cboLine, "Line", "Q_LINE_ALL");
                    LoadDataCbo(cboMachine, "Machine", "Q_MACHINE");
                    LoadDataCbo(cboStyle, "Style Name", "Q_STYLE_ALL");
                }
            }
        }

        private void cboLine_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                if (_tab == 0)
                {
                    LoadDataCbo(cboMachine, "Machine", "Q_MACHINE");
                    LoadDataCbo(cboStyle, "Style Name", "Q_STYLE");
                    LoadDataCbo(null, "Part", "Q_PART");
                }
                else if(_tab == 1)
                {
                    LoadDataCbo(cboMachine, "Machine", "Q_MACHINE");
                    LoadDataCbo(cboStyle, "Style Name", "Q_STYLE_ALL");
                }
            }
        }

        private void cboAssDate_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                LoadDataCbo(cboStyle, "Style Name", "Q_STYLE");
                LoadDataCbo(null, "Part", "Q_PART");
            }
        }

        private void cboStyle_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                if (_tab == 0)
                {
                    LoadDataCbo(null, "Part", "Q_PART");
                }
            }
        }


        private void cboWorkDate_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                if (_tab == 1)
                {
                    LoadDataCbo(cboStyle, "Style Name", "Q_STYLE_ALL");
                }
            }
        }

        private void cboMachine_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                if (_tab == 1)
                {
                    LoadDataCbo(cboStyle, "Style Name", "Q_STYLE_ALL");
                }
            }
        }

        private void gvwMain_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            try
            {
                if (grdMain.DataSource == null || gvwMain.RowCount < 1) return;

                if (e.Column.FieldName.ToString().Equals("CS_SIZE"))
                {
                    e.Appearance.BackColor = Color.LightYellow;
                }

                if (e.CellValue.ToString().ToUpper().Contains("SELECTED"))
                {
                    e.Appearance.BackColor = Color.FromArgb(248, 203, 173);
                }

                if (e.CellValue.ToString().ToUpper().Contains("SAVED"))
                {
                    e.Appearance.BackColor = Color.Green;
                    e.Appearance.ForeColor = Color.White;
                }
            }
            catch { }
        }

        private void gvwConfirm_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            try
            {
                if (grdConfirm.DataSource == null || gvwConfirm.RowCount < 1) return;

                if (e.Column.FieldName.ToString().Equals("CS_SIZE"))
                {
                    e.Appearance.BackColor = Color.LightYellow;
                }

                if (!e.Column.FieldName.ToString().Equals("LINE_NM") && !e.Column.FieldName.ToString().Equals("MLINE_CD"))
                {
                    if (gvwConfirm.GetRowCellValue(e.RowHandle, "CS_SIZE").ToString().ToUpper().Equals("TOTAL"))
                    {
                        e.Appearance.BackColor = Color.FromArgb(224, 255, 255);
                        e.Appearance.ForeColor = Color.Blue;
                    }
                }

                if (!e.Column.FieldName.ToString().Equals("LINE_NM"))
                {
                    if (gvwConfirm.GetRowCellValue(e.RowHandle, "MLINE_CD").ToString().ToUpper().Equals("S.TOTAL"))
                    {
                        e.Appearance.BackColor = Color.FromArgb(255, 219, 201);
                    }
                }
            }
            catch { }
        }

        private void gvwMain_MouseDown(object sender, MouseEventArgs e)
        {
            if (grdMain.DataSource == null || gvwMain.RowCount < 1) return;

            GridView view = sender as GridView;

            // Get the clicked cell information
            GridHitInfo hitInfo = view.CalcHitInfo(e.Location);
            hitInfoStart = null;

            // Check if the clicked area is within a valid cell
            if (hitInfo.InRow && hitInfo.RowHandle >= 0 && hitInfo.Column != null)
            {
                // Get the cell's row handle and column
                int rowHandle = hitInfo.RowHandle;
                int columnIndex = hitInfo.Column.VisibleIndex;
                string columnName = hitInfo.Column.FieldName;
                string _currentValue = view.GetRowCellValue(rowHandle, hitInfo.Column).ToString();

                if (columnIndex >= 6)
                {
                    if (!string.IsNullOrEmpty(_currentValue))
                    {
                        if (_currentValue.ToUpper().Contains("SAVED")) return;

                        _isMouseDown = true;
                        hitInfoStart = hitInfo;

                        string _line = gvwMain.GetRowCellValue(rowHandle, "LINE_CD").ToString();
                        string _mline = gvwMain.GetRowCellValue(rowHandle, "MLINE_CD").ToString();
                        string _hh = gvwMain.GetRowCellValue(rowHandle, "INPUT_PRIO").ToString();
                        string _style = gvwMain.GetRowCellValue(rowHandle, "STYLE_CD").ToString();
                        string _size = gvwMain.GetRowCellValue(rowHandle, "CS_SIZE").ToString();
                        string _part_no = columnName;
                        string _qty = gvwMain.GetRowCellValue(rowHandle, columnName).ToString();

                        bool _isDuplicate = false;

                        if (_dtSelected != null && _dtSelected.Rows.Count > 0)
                        {
                            for (int iRow = 0; iRow < _dtSelected.Rows.Count; iRow++)
                            {
                                if (_dtSelected.Rows[iRow]["LINE_CD"].ToString().Equals(_line) &&
                                    _dtSelected.Rows[iRow]["MLINE_CD"].ToString().Equals(_mline) &&
                                    _dtSelected.Rows[iRow]["INPUT_PRIO"].ToString().Equals(_hh) &&
                                    _dtSelected.Rows[iRow]["STYLE_CD"].ToString().Equals(_style) &&
                                    _dtSelected.Rows[iRow]["CS_SIZE"].ToString().Equals(_size) &&
                                    _dtSelected.Rows[iRow]["PFC_PART_NO"].ToString().Equals(_part_no))
                                {
                                    DataRow rowToDelete = _dtSelected.Rows[iRow];

                                    // Mark the row as deleted
                                    rowToDelete.Delete();

                                    // Accept the changes to reflect the deletion in the DataTable
                                    _dtSelected.AcceptChanges();
                                    _isDuplicate = true;
                                    _currentValue = _currentValue.Replace("_SELECTED", "");
                                    view.SetRowCellValue(rowHandle, hitInfo.Column, _currentValue);
                                    break;
                                }
                            }
                        }

                        if (!_isDuplicate)
                        {
                            // insert row values
                            _dtSelected.Rows.Add(new Object[]{
                                _line,
                                _mline,
                                _hh,
                                _style,
                                _size,
                                _part_no,
                                _qty
                            });

                            // Change the background color of the clicked cell
                            view.SetRowCellValue(rowHandle, hitInfo.Column, _currentValue + "_SELECTED");
                        }


                        view.LayoutChanged();
                    }
                }
            }
        }

        private void gvwMain_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            try
            {
                if (grdMain.DataSource == null || gvwMain.RowCount < 1) return;

                if (e.CellValue.ToString().ToUpper().Contains("SELECTED"))
                {
                    e.DisplayText = (e.CellValue.ToString()).Replace("_SELECTED", "");
                }

                if (e.CellValue.ToString().ToUpper().Contains("SAVED"))
                {
                    e.DisplayText = (e.CellValue.ToString()).Replace("_SAVED", "");
                }
            }
            catch { }
        }

        private void gvwConfirm_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            try
            {
                if (grdConfirm.DataSource == null || gvwConfirm.RowCount < 1) return;

                if (e.Column.FieldName.ToString().Equals("CS_SIZE"))
                {
                    if (e.CellValue.ToString().ToUpper().Contains("TOTAL"))
                    {
                        e.DisplayText = "";
                    }
                }
            }
            catch { }
        }

        public bool SaveData(string _type)
        {
            JPlatform.Client.CSIGMESBaseform6.frmSplashScreenWait frmSplash = new JPlatform.Client.CSIGMESBaseform6.frmSplashScreenWait();
            try
            {
                bool _result = true;
                DataTable dtData = null;
                DataTable dtData1 = null;

                P_MSPD90226A_S proc = new P_MSPD90226A_S();
                //string machineName = $"{SessionInfo.UserName}|{ Environment.MachineName}|{GetIPAddress()}";
                string machineName = $"{SessionInfo.UserName}";
                int iUpdate = 0, iCount = 0;
                frmSplash.Show();
                if (_type == "Q_SAVE")
                {
                    for (int iRow = 0; iRow < _dtSelected.Rows.Count; iRow++)
                    {
                        iUpdate++;
                        dtData = proc.SetParamData(dtData,
                                                  _type,
                                                  cboFactory.EditValue.ToString(),
                                                  cboPlant.EditValue.ToString(),
                                                  cboLine.EditValue.ToString(),
                                                  _dtSelected.Rows[iRow]["MLINE_CD"].ToString(),
                                                  cboWorkDate.yyyymmdd,
                                                  cboAssDate.yyyymmdd,
                                                  cboMachine.EditValue.ToString(),
                                                  _dtSelected.Rows[iRow]["INPUT_PRIO"].ToString(),
                                                  _dtSelected.Rows[iRow]["STYLE_CD"].ToString().Replace("-", ""),
                                                  _dtSelected.Rows[iRow]["PFC_PART_NO"].ToString(),
                                                  _dtSelected.Rows[iRow]["CS_SIZE"].ToString(),
                                                  _dtSelected.Rows[iRow]["DIR_QTY"].ToString(),
                                                  machineName,
                                                  "CSI.GMES.PD.MSPD90226A_S");

                        if (CommonProcessSave(dtData, proc.ProcName, proc.GetParamInfo(), grdMain))
                        {
                            dtData = null;
                            iCount++;
                        }
                        else
                        {
                            // break;
                        }
                    }
                  

                    if (iUpdate == iCount)
                    {
                        _result = true;
                    }
                    else
                    {
                        _result = false;
                    }
                }
                else if (_type == "Q_DELETE")
                {
                    dtData1 = proc.SetParamData(dtData1,
                                                  _type,
                                                  cboFactory.EditValue.ToString(),
                                                  cboPlant.EditValue.ToString(),
                                                  cboLine.EditValue.ToString(),
                                                   "",
                                                  cboWorkDate.yyyymmdd,
                                                  cboAssDate.yyyymmdd,
                                                  cboMachine.EditValue.ToString(),
                                                  "",
                                                  "",
                                                  "",
                                                  "",
                                                  "",
                                                  machineName,
                                                 "CSI.GMES.PD.MSPD90226A_S");

                    if (CommonProcessSave(dtData1, proc.ProcName, proc.GetParamInfo(), grdMain))
                    {
                        dtData1 = null;
                        _result = true;
                    }
                    else
                    {
                        _result = false;
                    }    
                   
                }    

                frmSplash.Close();
                return _result;
            }
            catch (Exception ex)
            {
                frmSplash.Close();
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        private void gvwMain_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                _isMouseDown = false;

                GridView view = sender as GridView;

                // Get the clicked cell information
                GridHitInfo hitInfo = view.CalcHitInfo(e.Location);

                // Check if the clicked area is within a valid cell
                if (hitInfo.InRow && hitInfo.RowHandle >= 0 && hitInfo.Column != null)
                {
                    // Get the cell's row handle and column
                    int rowHandle = hitInfo.RowHandle;
                    int columnIndex = hitInfo.Column.VisibleIndex;
                    string columnName = hitInfo.Column.FieldName;
                    string _currentValue = view.GetRowCellValue(rowHandle, hitInfo.Column).ToString();

                }
            }
            catch { }
        }

        private void gvwMain_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (_isMouseDown)
                {
                    GridView view = sender as GridView;
                    GridHitInfo hitInfo = view.CalcHitInfo(e.Location);

                    // Check if the clicked area is within a valid cell
                    if (hitInfo.InRow && hitInfo.RowHandle >= 0 && hitInfo.Column != null && hitInfo.InRowCell)
                    {
                        // Get the cell's row handle and column
                        int rowHandle = hitInfo.RowHandle;
                        int columnIndex = hitInfo.Column.VisibleIndex;
                        string columnName = hitInfo.Column.FieldName;
                        string _currentValue = view.GetRowCellValue(rowHandle, hitInfo.Column).ToString();

                        if (columnIndex >= 6)
                        {
                            if (!string.IsNullOrEmpty(_currentValue))
                            {
                                if (_currentValue.ToUpper().Contains("SAVED") || _currentValue.ToUpper().Contains("SELECTED")) return;

                                string _line = gvwMain.GetRowCellValue(rowHandle, "LINE_CD").ToString();
                                string _mline = gvwMain.GetRowCellValue(rowHandle, "MLINE_CD").ToString();
                                string _hh = gvwMain.GetRowCellValue(rowHandle, "INPUT_PRIO").ToString();
                                string _style = gvwMain.GetRowCellValue(rowHandle, "STYLE_CD").ToString();
                                string _size = gvwMain.GetRowCellValue(rowHandle, "CS_SIZE").ToString();
                                string _part_no = columnName;
                                string _qty = gvwMain.GetRowCellValue(rowHandle, columnName).ToString();

                                _dtSelected.Rows.Add(new Object[]{
                                    _line,
                                    _mline,
                                    _hh,
                                    _style,
                                    _size,
                                    _part_no,
                                    _qty
                                });

                                view.SetRowCellValue(rowHandle, hitInfo.Column, _currentValue + "_SELECTED");
                                view.LayoutChanged();
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void gvwMain_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            try
            {
                if (grdMain.DataSource == null || gvwMain.RowCount < 1) return;
            }
            catch { }
        }

        private void gvwMain_CustomDrawBandHeader(object sender, DevExpress.XtraGrid.Views.BandedGrid.BandHeaderCustomDrawEventArgs e)
        {
            if (e.Band == null) return;
            if (e.Band.AppearanceHeader.BackColor != Color.Empty)
            {
                e.Info.AllowColoring = true;
            }
        }

        private void tabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            _tab = tabControl.SelectedTabPageIndex;

            if(_tab == 1)
            {
                chkcboPart.Visible = false;
                lbPart.Visible = false;
                cboStyle.Width = 329;
                lblSave.Visible = false;
                lblSelect.Visible = false;
                lblAssDate.Visible = false;
                cboAssDate.Visible = false;

                cboHMS.Location = new Point(286, 37);
                lblHMS.Location = new Point(222, 37);

                lblStyle.Location = new Point(22, 69);
                cboStyle.Location = new Point(87, 69);

                lblPlant.Location = new Point(237, 7);
                cboPlant.Location = new Point(286, 7);
                lblLine.Location = new Point(449, 7);
                cboLine.Location = new Point(490, 7);
                lblMachine.Location = new Point(419, 37);
                cboMachine.Location = new Point(490, 37);
            }
            else if (_tab == 0)
            {
                chkcboPart.Visible = true;
                lbPart.Visible = true;
                cboStyle.Width = 130;
                lblSave.Visible = true;
               // lblSelect.Visible = true;
                lblAssDate.Visible = true;
                cboAssDate.Visible = true;

                cboHMS.Location = new Point(87, 69);
                lblHMS.Location = new Point(22, 69);

                lblStyle.Location = new Point(223, 68);
                cboStyle.Location = new Point(326, 68);

                lblPlant.Location = new Point(277, 7);
                cboPlant.Location = new Point(326, 7);
                lblLine.Location = new Point(489, 7);
                cboLine.Location = new Point(530, 7);
                lblMachine.Location = new Point(459, 37);
                cboMachine.Location = new Point(530, 37);
            }

            _firstLoad = true;
            InitCombobox();
            _firstLoad = false;
        }

        #endregion

        #region Database

        public class P_MSPD90226A_Q : BaseProcClass
        {
            public P_MSPD90226A_Q()
            {
                // Modify Code : Procedure Name
                _ProcName = "LMES.P_MSPD90226A_Q";
                ParamAdd();
            }
            private void ParamAdd()
            {
                _ParamInfo.Add(new ParamInfo("@ARG_WORK_TYPE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_PLANT", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_LINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_MLINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_STYLE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_DATE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_ASS_DATE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_MACHINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_HMS", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_PART", "Varchar", 0, "Input", typeof(System.String)));
            }
            public DataTable SetParamData(DataTable dataTable,
                                        System.String ARG_WORK_TYPE,
                                        System.String ARG_PLANT,
                                        System.String ARG_LINE,
                                        System.String ARG_MLINE,
                                        System.String ARG_STYLE,
                                        System.String ARG_DATE,
                                        System.String ARG_ASS_DATE,
                                        System.String ARG_MACHINE,
                                        System.String ARG_HMS,
                                        System.String ARG_PART)
            {
                if (dataTable == null)
                {
                    dataTable = new DataTable(_ProcName);
                    foreach (ParamInfo pi in _ParamInfo)
                    {
                        dataTable.Columns.Add(pi.ParamName, pi.TypeClass);
                    }
                }
                // Modify Code : Procedure Parameter
                object[] objData = new object[] {
                                                ARG_WORK_TYPE,
                                                ARG_PLANT,
                                                ARG_LINE,
                                                ARG_MLINE,
                                                ARG_STYLE,
                                                ARG_DATE,
                                                ARG_ASS_DATE,
                                                ARG_MACHINE,
                                                ARG_HMS,
                                                ARG_PART
                };
                dataTable.Rows.Add(objData);
                return dataTable;
            }
        }

        public class P_MSPD90226A_COMBO : BaseProcClass
        {
            public P_MSPD90226A_COMBO()
            {
                // Modify Code : Procedure Name
                _ProcName = "LMES.P_MSPD90226A_COMBO";
                ParamAdd();
            }
            private void ParamAdd()
            {
                _ParamInfo.Add(new ParamInfo("@ARG_WORK_TYPE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_PLANT", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_LINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_MLINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_STYLE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_DATE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_ASS_DATE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_HMS", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_MACHINE", "Varchar", 100, "Input", typeof(System.String)));
            }
            public DataTable SetParamData(DataTable dataTable,
                                        System.String ARG_WORK_TYPE,
                                        System.String ARG_PLANT,
                                        System.String ARG_LINE,
                                        System.String ARG_MLINE,
                                        System.String ARG_STYLE,
                                        System.String ARG_DATE,
                                        System.String ARG_ASS_DATE,
                                        System.String ARG_HMS,
                                        System.String ARG_MACHINE)
            {
                if (dataTable == null)
                {
                    dataTable = new DataTable(_ProcName);
                    foreach (ParamInfo pi in _ParamInfo)
                    {
                        dataTable.Columns.Add(pi.ParamName, pi.TypeClass);
                    }
                }
                // Modify Code : Procedure Parameter
                object[] objData = new object[] {
                    ARG_WORK_TYPE,
                    ARG_PLANT,
                    ARG_LINE,
                    ARG_MLINE,
                    ARG_STYLE,
                    ARG_DATE,
                    ARG_ASS_DATE,
                    ARG_HMS,
                    ARG_MACHINE
                };
                dataTable.Rows.Add(objData);
                return dataTable;
            }
        }

        public class P_MSPD90226A_S : BaseProcClass
        {
            public P_MSPD90226A_S()
            {
                // Modify Code : Procedure Name
                _ProcName = "LMES.P_MSPD90226A_S";
                ParamAdd();
            }
            private void ParamAdd()
            {
                _ParamInfo.Add(new ParamInfo("@ARG_TYPE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_PLANT", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_LINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_MLINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_STIT_LINE", "Varchar2", 100, "Input", typeof(System.String)));

                _ParamInfo.Add(new ParamInfo("@ARG_WO_YMD", "Varchar2", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_ASY_YMD", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_MACHINE_ID", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_HH", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_STYLE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_PFC_PART_NO", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_CS_SIZE", "Varchar2", 0, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_QTY", "Varchar", 100, "Input", typeof(System.String)));

                _ParamInfo.Add(new ParamInfo("@ARG_CREATE_PC", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_CREATE_PROGRAM_ID", "Varchar", 100, "Input", typeof(System.String)));
            }
            public DataTable SetParamData(DataTable dataTable,
                                        System.String ARG_TYPE,
                                        System.String ARG_PLANT,
                                        System.String ARG_LINE,
                                        System.String ARG_MLINE,
                                        System.String ARG_STIT_LINE,

                                        System.String ARG_WO_YMD,
                                        System.String ARG_ASY_YMD,
                                        System.String ARG_MACHINE_ID,
                                        System.String ARG_HH,
                                        System.String ARG_STYLE,
                                        System.String ARG_PFC_PART_NO,
                                        System.String ARG_CS_SIZE,
                                        System.String ARG_QTY,

                                        System.String ARG_CREATE_PC,
                                        System.String ARG_CREATE_PROGRAM_ID)
            {
                if (dataTable == null)
                {
                    dataTable = new DataTable(_ProcName);
                    foreach (ParamInfo pi in _ParamInfo)
                    {
                        dataTable.Columns.Add(pi.ParamName, pi.TypeClass);
                    }
                }
                // Modify Code : Procedure Parameter
                object[] objData = new object[] {
                    ARG_TYPE,
                    ARG_PLANT,
                    ARG_LINE,
                    ARG_MLINE,
                    ARG_STIT_LINE,
                    ARG_WO_YMD,
                    ARG_ASY_YMD,
                    ARG_MACHINE_ID,
                    ARG_HH,
                    ARG_STYLE,
                    ARG_PFC_PART_NO,
                    ARG_CS_SIZE,
                    ARG_QTY,
                    ARG_CREATE_PC,
                    ARG_CREATE_PROGRAM_ID
                };
                dataTable.Rows.Add(objData);
                return dataTable;
            }
        }

        #endregion

        DataTable GetDataTable(GridView view)
        {
            DataTable dt = new DataTable();
            foreach (GridColumn c in view.Columns)
                dt.Columns.Add(c.FieldName, c.ColumnType);
            for (int r = 0; r < view.RowCount; r++)
            {
                object[] rowValues = new object[dt.Columns.Count];
                for (int c = 0; c < dt.Columns.Count; c++)
                    rowValues[c] = view.GetRowCellValue(r, dt.Columns[c].ColumnName);
                dt.Rows.Add(rowValues);
            }
            return dt;
        }

        private DataTable LINQResultToDataTable<T>(IEnumerable<T> Linqlist)
        {
            DataTable dt = new DataTable();
            PropertyInfo[] columns = null;
            if (Linqlist == null) return dt;
            foreach (T Record in Linqlist)
            {
                if (columns == null)
                {
                    columns = ((Type)Record.GetType()).GetProperties();
                    foreach (PropertyInfo GetProperty in columns)
                    {
                        Type colType = GetProperty.PropertyType;

                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition()
                        == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }

                        dt.Columns.Add(new DataColumn(GetProperty.Name, colType));
                    }
                }
                DataRow dr = dt.NewRow();
                foreach (PropertyInfo pinfo in columns)
                {
                    dr[pinfo.Name] = pinfo.GetValue(Record, null) == null ? DBNull.Value : pinfo.GetValue
                    (Record, null);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}