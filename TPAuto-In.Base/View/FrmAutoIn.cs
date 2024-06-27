using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TP_Automation.Logger.Interface;
using TP_Automation.Logger.Model;
using TPAuto_In.Base.Configuration;

namespace TPAuto_In.Base.View
{
    public partial class FrmAutoIn : Form
    {
        bool expanded = true;
        private DateTime DataInicio;

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,     // x-coordinate of upper-left corner
            int nTopRect,      // y-coordinate of upper-left corner
            int nRightRect,    // x-coordinate of lower-right corner
            int nBottomRect,   // y-coordinate of lower-right corner
            int nWidthEllipse, // height of ellipse
            int nHeightEllipse // width of ellipse
        );

        public FrmAutoIn()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
        }

        public FrmAutoIn(DateTime dataInicio)
        {
            InitializeComponent();
            this.DataInicio = dataInicio;
            this.picFechar.Visible = false;
            Application.DoEvents();
        }

        public void LiberaEncerramento()
        {
            this.picFechar.Visible = true;
            Application.DoEvents();
        }

        public void TrayMessage(string message)
        {
            notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
            notifyIcon.Icon = SystemIcons.Exclamation;
            notifyIcon.Visible = true;
            notifyIcon.BalloonTipText = message;
            notifyIcon.BalloonTipTitle = "TPAuto-in";
            notifyIcon.ShowBalloonTip(5000);
        }


        private void FrmAutoIn_Click(object sender, EventArgs e)
        {

        }

        private void FrmAutoIn_DoubleClick(object sender, EventArgs e)
        {

        }

        private void FrmAutoIn_Load(object sender, EventArgs e)
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fvi.FileVersion;
            lblVersao.Text = lblVersao.Text.Replace("{{VERSAO}}", version);
            this.TopMost = true;
        }

        private void FrmAutoIn_Paint(object sender, PaintEventArgs e)
        {

        }

        private void picFechar_Click(object sender, EventArgs e)
        {
            try
            {
                var dialog = MessageBox.Show("Será registrado que houve um encerramento manual, tem certeza que deseja encerrar o TP Auto-IN?", "TP Auto-IN", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialog == DialogResult.Yes)
                {
                    GravarLogFechamentoManual();
                    Application.Exit();
                    Environment.Exit(0);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro ao encerrar o processo manualmente");
            }
        }

        private void GravarLogFechamentoManual()
        {
            try
            {
                var log = new Log();
                LogData logData = new LogData();
                logData.LogsProcessData = new List<LogProcessData>();

                logData.StartDate = DataInicio;
                logData.Hostname = Environment.MachineName;
                logData.Username = Environment.UserName;
                logData.AutomationID = Global.AutomationId;
                logData.ClientID = Global.ClientId;

                var logEncerramento = new LogProcessData();
                logEncerramento.StartDate = DateTime.Now;
                logEncerramento.IdProcessAutomation = Global.IdProcessoEncerramentoManual;

                logEncerramento.EndDate = DateTime.Now;
                logEncerramento.Success = true;
                logEncerramento.Message = "Automação encerrada manualmente";
                logData.LogsProcessData.Add(logEncerramento);

                logData.Success = false;
                logData.Message = "Automação encerrada manualmente";
                logData.EndDate = DateTime.Now;

                var result = log.SaveLog(logData, Global.IsProd);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void FrmAutoIn_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
            Environment.Exit(0);
        }

        private void panel1_DoubleClick(object sender, EventArgs e)
        {
            Application.DoEvents();
        }
    }
}
