using BLL;
using DevExpress.XtraCharts.Designer.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrintingLinks;
using DTO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace GUI.ChildrenForms.HoaDonChiTieu.DanhSachHoaDon
{
    public partial class fChiTietHoaDon : DevExpress.XtraEditors.XtraForm
    {
        private DTO_HOADON hoaDon;
        private SimpleButton btn_xuatHoaDon;
        public fChiTietHoaDon(DTO_HOADON hoaDon)
        {
            InitializeComponent();
            this.hoaDon = hoaDon;

            CreatePrintButtons();
        }
        private void CreatePrintButtons()
        {
            btn_xuatHoaDon = new SimpleButton();
            btn_xuatHoaDon.Text = "Xuất hóa đơn";
            btn_xuatHoaDon.Size = new Size(160, 50);
            btn_xuatHoaDon.Location =
            new Point(
                    this.ClientSize.Width - 200,
                    this.ClientSize.Height - 65
                );
            btn_xuatHoaDon.Anchor =
                AnchorStyles.Bottom | AnchorStyles.Right;
            btn_xuatHoaDon.Click += Btn_xuatHoaDon_Click;
            this.Controls.Add(btn_xuatHoaDon);
            btn_xuatHoaDon.BringToFront();
        }
        private void Btn_xuatHoaDon_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text File|*.txt";
            sfd.FileName =
                "HoaDon_" +
                hoaDon.TenPhong +
                "_" +
                hoaDon.Thang +
                "_" +
                hoaDon.Nam;
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                StringBuilder hd = new StringBuilder();
                hd.AppendLine("====================================");
                hd.AppendLine("          HOA DON PHONG TRO         ");
                hd.AppendLine("====================================");
                hd.AppendLine("");
                hd.AppendLine("Phong: " + hoaDon.TenPhong);
                hd.AppendLine("Thang: " +
                    hoaDon.Thang + "/" +
                    hoaDon.Nam);
                hd.AppendLine("");
                hd.AppendLine("Tien phong : " +
                    String.Format("{0:#,##0}", hoaDon.TienPhong));
                hd.AppendLine("Tien dien  : " +
                    String.Format("{0:#,##0}", hoaDon.TienDien));
                hd.AppendLine("Tien nuoc  : " +
                    String.Format("{0:#,##0}", hoaDon.TienNuoc));
                hd.AppendLine("Tien dich vu: " +
                    String.Format("{0:#,##0}", hoaDon.TienDichVu));
                hd.AppendLine("");
                hd.AppendLine("------------------------------------");
                hd.AppendLine("TONG TIEN: " +
                    String.Format("{0:#,##0} đ", hoaDon.TongTien));
                hd.AppendLine("------------------------------------");
                hd.AppendLine("");
                hd.AppendLine(
                    Reader.ReadNumber(
                        (decimal)hoaDon.TongTien
                    )
                );
                hd.AppendLine("");
                hd.AppendLine(
                    hoaDon.DaThanhToan == "True"
                    ? "TRANG THAI: DA THANH TOAN"
                    : "TRANG THAI: CHUA THANH TOAN"
                );
                File.WriteAllText(
                    sfd.FileName,
                    hd.ToString(),
                    Encoding.UTF8
                );
                XtraMessageBox.Show(
                    "Xuất hóa đơn thành công!",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
        }

        private void SetMonthYear()
        {
            int thang = this.hoaDon.Thang;
            int nam = this.hoaDon.Nam;

            this.lb_thangNam.Text = String.Format("Tháng {0:d2} - {1}", thang, nam);
        }

        private void SetServices()
        {
            DataTable dt = BLL_HOADON_DICHVU.LayDanhSachDichVuTheoMaHoaDon(this.hoaDon);
            this.gridControl_dichVuPhong.DataSource = dt;

            this.gridView_dichVuPhong.Columns["MAHOADON"].Visible = false;
            this.gridView_dichVuPhong.Columns["MADICHVU"].Visible = false;
            this.gridView_dichVuPhong.Columns["TENDICHVU"].Caption = "Tên dịch vụ";
            this.gridView_dichVuPhong.Columns["GIADICHVU"].Caption = "Giá dịch vụ";

            this.lb_tongTienDichVu.Text = String.Format("{0:#,##0}", this.hoaDon.TienDichVu);
        }

        private void btn_xong_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void fChiTietHoaDon_Load(object sender, EventArgs e)
        {
            SetMonthYear();
            this.lb_tenPhong.Text = this.hoaDon.TenPhong;
            this.lb_giaPhong.Text = String.Format("{0:#,##0}", this.hoaDon.TienPhong);
            this.lb_tongTien.Text = String.Format("{0:#,##0} đ", this.hoaDon.TongTien);
            this.lb_tienBangChu.Text = Reader.ReadNumber((decimal)this.hoaDon.TongTien);


            this.lb_tongTienDien.Text = String.Format("{0:#,##0}", this.hoaDon.TienDien);
            this.lb_tongTienNuoc.Text = String.Format("{0:#,##0}", this.hoaDon.TienNuoc);

            DTO_HOADON_DIENNUOC hdnt = BLL_HOADON_DIENNUOC.LayThongTinThangTruoc(hoaDon);
            this.txt_soDienCu.Text = hdnt.SoDien.ToString();
            this.txt_soNuocCu.Text = hdnt.SoNuoc.ToString();

            DTO_HOADON_DIENNUOC hdnht = BLL_HOADON_DIENNUOC.LayThongTinThangHienTai(hoaDon);
            this.txt_soDienMoi.Text = hdnht.SoDien.ToString();
            this.txt_soNuocMoi.Text = hdnht.SoNuoc.ToString();
            this.lb_giaDien.Text = String.Format("{0:#,##0}", hdnht.GiaDien);
            this.lb_giaNuoc.Text = String.Format("{0:#,##0}", hdnht.GiaNuoc);

            this.txt_soDienTieuThu.Text = (hdnht.SoDien - hdnt.SoDien).ToString();
            this.txt_soNuocTieuThu.Text = (hdnht.SoNuoc - hdnt.SoNuoc).ToString();

            SetServices();
        }
    }
}