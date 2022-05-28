using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSDLPT_NGANHANG
{
    public partial class formAccount : Form
    {
        public formAccount()
        {
            InitializeComponent();
        }

        private void taiKhoanBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.bdsTK.EndEdit();
            this.tableAdapterManager.UpdateAll(this.dSm);

        }

        SqlConnection conn1 = new SqlConnection(Program.connstr_pulisher);
        SqlCommand cmd = new SqlCommand();
        int pos = 0;
        String macn = "";
        DateTime dt= DateTime.Today;

        private string cmd1 = "SELECT * FROM NGANHANG WHERE MACN = 'BENTHANH'";
        private string cmd2 = "SELECT * FROM NGANHANG WHERE MACN = 'TANDINH'";
        
        private void formAccount_Load(object sender, EventArgs e)
        {
            dSm.EnforceConstraints = false;
    
            cbbChinhanh.DataSource = Program.bds_dspm;
            cbbChinhanh.DisplayMember = "TENCN";
            cbbChinhanh.ValueMember = "TENSERVER";
            cbbChinhanh.SelectedIndex = Program.mChinhanh;
            
            if (Program.mChinhanh == 0)
            {
                SqlDataAdapter adapter = new SqlDataAdapter(cmd1, Program.connstr);
                adapter.Fill(this.dSm.TaiKhoan);
            }
            else
            {
                SqlDataAdapter adapter = new SqlDataAdapter(cmd2, Program.connstr);
                adapter.Fill(this.dSm.TaiKhoan);
            }

            this.gD_GOIRUTTableAdapter.Connection.ConnectionString = Program.connstr;
            this.gD_GOIRUTTableAdapter.Fill(this.dSm.GD_GOIRUT);

            btUdo.Enabled = btSave.Enabled = false;

            macn = ((DataRowView)bdsTK[0])["MACN"].ToString();
           
            if (Program.mGroup == "NGANHANG")
            {
                cbbChinhanh.Enabled = true;
                btAdd.Enabled = btDel.Enabled = btFix.Enabled =  btSave.Enabled= btUdo.Enabled= false;
                pnTT.Enabled = false;
            }
            else
            {
                cbbChinhanh.Enabled = false;
            }


        }

        private bool checkSOTK()
        {
            SqlConnection conn = new SqlConnection(Program.connstr_pulisher);
            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT COUNT(*) FROM TaiKhoan WHERE SOTK = @sotk", conn);
            cmd.Parameters.AddWithValue("@sotk", txtSOTK.Text);
            int UserExist = (int)cmd.ExecuteScalar();

            if (UserExist > 0)
            {
                conn.Close();
                return true;

            }
            conn.Close();
            return false;

        }

        private bool checkCMND()
        {
            SqlConnection conn = new SqlConnection(Program.connstr_pulisher);
            conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT COUNT(*) FROM KhachHang WHERE (CMND = @cmnd)", conn);
            cmd.Parameters.AddWithValue("@cmnd", txtCMND.Text);
            int UserExist = (int)cmd.ExecuteScalar();

            if (UserExist > 0)
            {
                conn.Close();
                return true;

            }
            conn.Close();
            return false;

        }


        private bool checkKH()
        {
            
            Program.conn.Open();
            SqlCommand cmd = new SqlCommand("SELECT COUNT(*) FROM KhachHang WHERE CMND = @cmnd", Program.conn);
            cmd.Parameters.AddWithValue("@cmnd", txtCMND.Text);
            int UserExist = (int)cmd.ExecuteScalar();

            if (UserExist > 0)
            {
                Program.conn.Close();
                return true;

            }
            Program.conn.Close();
            return false;

        }

        private void taiKhoanGridControl_Click(object sender, EventArgs e)
        {

        }

        private void btAdd_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            pos = bdsTK.Position;
            pnTT.Enabled = true;
            bdsTK.AddNew();
            dtNGAY.DateTime = dt;

            if (Program.mChinhanh == 0)
            {
                txtMACN.Text = "BENTHANH";
            }
            else
            {
                txtMACN.Text = "TANDINH";
            }

            btAdd.Enabled = btDel.Enabled = btFix.Enabled = btExit.Enabled = btReload.Enabled = false;
            btSave.Enabled =btUdo.Enabled = true;
            gcTK.Enabled = false;

        }


        private void btDel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (bdsGR.Count > 0)
            {
                MessageBox.Show("Không thể xóa số tài khoản này vì đã lập giao dịch.");
                return;
            }

            if (MessageBox.Show("Bạn muốn xóa số tài khoản này?", "Xác nhận", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                try
                {
                    string query = "DELETE FROM TaiKhoan WHERE SOTK = @SOTK";
                    Program.conn.Open();

                    cmd = new SqlCommand(query, Program.conn);
                    cmd.Parameters.AddWithValue("@SOTK", txtSOTK.Text);

                    cmd.ExecuteNonQuery();
                    this.taiKhoanTableAdapter.Connection.ConnectionString = Program.connstr;
                    this.taiKhoanTableAdapter.Fill(this.dSm.TaiKhoan);

                    MessageBox.Show("Xóa thành công.");
                    Program.conn.Close();

                }
                catch (Exception ex)
                {
                    Program.conn.Close();
                    return;

                }
            }

            if (bdsTK.Count == 0) btDel.Enabled = false;
        }

        private void btFix_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

            if (txtSOTK.Text == "")
            {
                MessageBox.Show("Số tài khoản không được để trống.");
                txtSOTK.Focus();
                return;

            }
            else if (txtCMND.Text == "")
            {
                MessageBox.Show("CMND không được để trống.");
                txtCMND.Focus();
                return;

            }
            else if (txtSODU.Text == "")
            {
                MessageBox.Show("Số dư không được để trống.");
                txtSODU.Focus();
                return;

            }
            else if (dtNGAY.Text == "")
            {
                MessageBox.Show("Ngày mở tài khoản không được để trống.");
                dtNGAY.Focus();
                return;

            }        
            else if (txtMACN.Text == "")
            {
                MessageBox.Show("Mã chi nhánh không được để trống.");
                txtMACN.Focus();
                return;

            }
           

            string sotk = ((DataRowView)bdsTK[bdsTK.Position])["SOTK"].ToString();

            if (txtSOTK.Text != sotk)
            {
                MessageBox.Show("Số tài khoản cần sửa không trùng.");
            }
            else
            {
                try
                {
                    string query = "UPDATE TaiKhoan SET CMND=@CMND, SODU=@SODU, MACN=@MACN, NGAYMOTK=@NGAYMOTK WHERE SOTK=@SOTK";

                    Program.conn.Open();
                    cmd = new SqlCommand(query, Program.conn);
                    cmd = new SqlCommand(query, Program.conn);
                    cmd.Parameters.AddWithValue("@SOTK", txtSOTK.Text);
                    cmd.Parameters.AddWithValue("@CMND", txtCMND.Text);
                    cmd.Parameters.AddWithValue("@SODU", txtSODU.Text);
                    cmd.Parameters.AddWithValue("@MACN", txtMACN.Text);
                    cmd.Parameters.AddWithValue("@NGAYMOTK", dtNGAY.DateTime);


                    cmd.ExecuteNonQuery();
                    this.taiKhoanTableAdapter.Connection.ConnectionString = Program.connstr;
                    this.taiKhoanTableAdapter.Fill(this.dSm.TaiKhoan);

                    MessageBox.Show("Sửa thành công.");
                    Program.conn.Close();

                }
                catch (Exception ex)
                {
                    Program.conn.Close();
                    return;
                }
            }
        }

        private void btReload_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.taiKhoanTableAdapter.Fill(this.dSm.TaiKhoan);
        }

 

        private void cbbChinhanh_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbbChinhanh.SelectedValue.ToString() == "System.Data.DataRowView") return;

            Program.servername = cbbChinhanh.SelectedValue.ToString();

            if (cbbChinhanh.SelectedIndex != Program.mChinhanh)
            {
                Program.mlogin = Program.remotelogin;
                Program.password = Program.remotepassword;

            }
            else
            {
                Program.mlogin = Program.mloginDN;
                Program.password = Program.passwordDN;
            }

            if (Program.KetNoi() == 0)
            {
                MessageBox.Show("Lỗi kết nối!");
            }
            else
            {
                if (Program.mChinhanh == 0)
                {
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd1, Program.connstr);
                    dSm.Clear();
                    adapter.Fill(this.dSm.TaiKhoan);
                }
                else
                {
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd2, Program.connstr);
                    dSm.Clear();
                    adapter.Fill(this.dSm.TaiKhoan);
                }

                this.gD_GOIRUTTableAdapter.Connection.ConnectionString = Program.connstr;
                this.gD_GOIRUTTableAdapter.Fill(this.dSm.GD_GOIRUT);
            }

        }

        private void btSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            int x = Int32.Parse(txtSOTK.Text);
            if (txtSOTK.Text == "")
            {
                MessageBox.Show("Số tài khoản không được để trống.");
                txtSOTK.Focus();
                return;

            }else if (txtSOTK.Text.Length > 5 || txtSOTK.Text.Length < 4)
            {
                MessageBox.Show("Số tài khoản phải có 4 chữ số.");
                txtSOTK.Focus();
                return;
            }
            else if (txtCMND.Text == "")
            {
                MessageBox.Show("CMND không được để trống.");
                txtCMND.Focus();
                return;

            }
            else if (txtSODU.Text == "")
            {
                MessageBox.Show("Số dư không được để trống.");
                txtSODU.Focus();
                return;

            }
            else if (dtNGAY.Text == "")
            {
                MessageBox.Show("Ngày mở tài khoản không được để trống.");
                dtNGAY.Focus();
                return;

            }
            else if (txtMACN.Text == "")
            {
                MessageBox.Show("Mã chi nhánh không được để trống.");
                txtMACN.Focus();
                return;
            }

            if (checkCMND() == false)
            {
                MessageBox.Show("CMND này không tồn tại.");
            }


            if (checkSOTK() == true)
            {
                MessageBox.Show("Số tài khoản này đã tồn tại.");
            }
            else
            {
                if (checkKH() == true)
                {
                    try
                    {

                        string query = "INSERT INTO TaiKhoan(SOTK, CMND, SODU, MACN,NGAYMOTK)" +
                            "VALUES(@SOTK, @CMND, @SODU, @MACN,@NGAYMOTK)";


                        Program.conn.Open();
                        cmd = new SqlCommand(query, Program.conn);
                        cmd.Parameters.AddWithValue("@SOTK", txtSOTK.Text);
                        cmd.Parameters.AddWithValue("@CMND", txtCMND.Text);
                        cmd.Parameters.AddWithValue("@SODU", txtSODU.Text);
                        cmd.Parameters.AddWithValue("@MACN", txtMACN.Text);
                        cmd.Parameters.AddWithValue("@NGAYMOTK", dtNGAY.DateTime);


                        cmd.ExecuteNonQuery();
                        this.taiKhoanTableAdapter.Connection.ConnectionString = Program.connstr;
                        this.taiKhoanTableAdapter.Fill(this.dSm.TaiKhoan);

                        MessageBox.Show("Mở tài khoản thành công.");
                        Program.conn.Close();


                        gcTK.Enabled = true;
                        btAdd.Enabled = btDel.Enabled = btFix.Enabled = btReload.Enabled = btExit.Enabled = true;
                        btSave.Enabled = btUdo.Enabled = false;

                    }
                    catch (Exception ex)
                    {
                        Program.conn.Close();
                        return;
                    }
                }
                else
                {
                    try
                    {

                        string query = "INSERT INTO TaiKhoan(SOTK, CMND, SODU, MACN,NGAYMOTK)" +
                            "VALUES(@SOTK, @CMND, @SODU, @MACN,@NGAYMOTK)";


                        conn1.Open();
                        cmd = new SqlCommand(query, conn1);
                        cmd.Parameters.AddWithValue("@SOTK", txtSOTK.Text);
                        cmd.Parameters.AddWithValue("@CMND", txtCMND.Text);
                        cmd.Parameters.AddWithValue("@SODU", txtSODU.Text);
                        cmd.Parameters.AddWithValue("@MACN", txtMACN.Text);
                        cmd.Parameters.AddWithValue("@NGAYMOTK", dtNGAY.DateTime);


                        cmd.ExecuteNonQuery();
                        this.taiKhoanTableAdapter.Connection.ConnectionString = Program.connstr;
                        this.taiKhoanTableAdapter.Fill(this.dSm.TaiKhoan);

                        MessageBox.Show("Mở tài khoản thành công.");
                        conn1.Close();

                        gcTK.Enabled = true;
                        btAdd.Enabled = btDel.Enabled = btFix.Enabled = btReload.Enabled = btExit.Enabled = true;
                        btSave.Enabled = btUdo.Enabled = false;

                    }
                    catch (Exception ex)
                    {
                        conn1.Close();
                        return;
                    }
                }
            }
       

         }

        private void btUdo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            bdsTK.CancelEdit();
            if (btAdd.Enabled == false) bdsTK.Position = pos;
            gcTK.Enabled = true;
            //pnTT.Enabled = false;
            btAdd.Enabled = btDel.Enabled = btFix.Enabled = btReload.Enabled = btExit.Enabled = true;
            btSave.Enabled=btUdo.Enabled = false;

        }

        private void btExit_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Close();
        }
    }
}
