using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ERP_PR
{
    public partial class Choose_Little_thing : Form
    {
        //類別
        SelectData selectdata = new SelectData();

        //SheetView[]
        SheetView[] sv_array;

        //CellType
        ButtonCellType btn = new ButtonCellType();
        CheckBoxCellType ckb = new CheckBoxCellType();

        //TextBox
        public TextBox txb,txb_code;

        //語言
        public ToolStripComboBox lg = new ToolStripComboBox();

        //排序&切換語言會用到的陣列
        public string[] SetName_B;
        public string[] tag_list_B;

        //一般宣告
        public string[] invisible_col=new string[] { };   //隱藏欄位
        public string DataCol = "", DataCodeCol="";
        
        //初始化
        public Choose_Little_thing()
        {
            InitializeComponent();

            //設定連線
            selectdata.GetDBInf();

            //SheetView集合
            sv_array = new SheetView[] { sv_B };

            //語言設定
            lg.TextChanged += new EventHandler(Lg_change);
        }

        //載入
        private void Choose_Little_thing_Load(object sender, EventArgs e)
        {
            //取得表身小標tag list
            tag_list_B = selectdata.select_small_tag_list(SetName_B);

            //設定標題
            selectdata.Set_Title(sv_B, SetName_B, null, "", 0, null);

            //合併list並查詢所有語言
            List<string> merge_list = new List<string>();
            merge_list.AddRange(tag_list_B);
            selectdata.set_All_lg(merge_list.Distinct().ToList());

            //設定欄位
            Set_Col();

            //撈資料
            Select_Data();

            //切換語言
            Change_Language();
        }

        //設定欄位
        public void Set_Col()
        {
            //隱藏欄位
            foreach (string item in invisible_col)
                sv_B.Columns[item].Visible = false;
        }

        //******************************篩選******************************

        //透過TextBox欄位做篩選
        private void txb_Search_TextChanged(object sender, EventArgs e)
        {
            //撈資料
            Select_Data();
        }

        //撈資料
        public void Select_Data()
        {
            sv_B.DataSource = selectdata.select_little_thing(SetName_B, txb_Search.Text.Trim());
            if (sv_B.RowCount == 0)
                MessageBox.Show("查無資料");
            //設定欄寬
            Call_Set_ColWidth();
        }

        //******************************篩選******************************

        //切換語言事件
        private void Lg_change(object sender, EventArgs e)
        {
            Change_Language();
        }

        //語言轉換
        public void Change_Language()
        {
            selectdata.set_sp_language_NEW(lg.Text, sv_B, tag_list_B, null);   //表身
        }

        //呼叫 "設定欄寬符合資料長度" 方法
        public void Call_Set_ColWidth()
        {
            //設定欄寬符合資料長度
            selectdata.Set_ColWidth_Fit_DataWidth(sv_array);

            ////設定sp寬
            //Size size = new Size((sv_B.Columns.Count-invisible_col.Count()) * 115,200);
            //this.Size = size;
            //設定sp寬
            int width = 0;
            for (int i = 0; i < sv_B.Columns.Count; i++)
                if (!Array.Exists(invisible_col, element => element == sv_B.Columns[i].Tag.ToString()))
                    width += Convert.ToInt32(sv_B.Columns[i].Width);
            //Size size = new Size(((sv_B.Columns.Count-invisible_col.Count()) * 120)>300 ? 300: ((sv_B.Columns.Count - invisible_col.Count()) * 120), 150);
            Size size = new Size(width + 20, 150);
            this.Size = size;
        }
        //當進入某Cell
        private void fpSpread1_EnterCell(object sender, EnterCellEventArgs e)
        {
            //sv_B.Rows[e.Row].BackColor = Color.Green;
        }

        //設定按ESC取消
        private void Choose_Little_thing_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                this.Close();
        }

        //雙擊某列 取得對應編號?
        private void fpSpread1_CellDoubleClick(object sender, CellClickEventArgs e)
        {
            txb.Text = sv_B.Cells[e.Row,sv_B.Columns[DataCol].Index].Text;
            this.Close();
            //if (!string.IsNullOrEmpty(DataCodeCol))
            //{
            //    this.Owner.Tag = sv_B.Cells[e.Row, sv_B.Columns[DataCol].Index].Text;
            //}
            ////id
        }
    }
}
