using System;
using System.Windows;

namespace KMIQ.Form
{
    /// <summary>
    /// frmProgress.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class frmProgress : Window
    {
        public frmProgress()
        {
            InitializeComponent();
            Body.MouseLeftButtonDown += (o, e) => { try {  DragMove(); } catch { } };
        }

        internal void initializeProgress(int maxValue)
        {
            lbl_Progress.Text = "";
            lbl_Status.Text = "";
            Progress.Maximum = maxValue;
            Progress.Value = 0;
        }

        internal void updateProgress2(bool imi, string status)
        {
            lbl_Progress.Text = status;
            lbl_Status.Text = "";
            Progress.IsIndeterminate = true;
            Progress.Value = 0;
            System.Windows.Forms.Application.DoEvents();
        }

        internal void updateProgress(int value, string status)
        {
            lbl_Status.Text = status;
            lbl_Progress.Text = string.Format("{0:N1}%", Convert.ToDecimal(value) / Convert.ToDecimal(Progress.Maximum) * 100);
            Progress.Value = value;
            Progress.IsIndeterminate = false;
            System.Windows.Forms.Application.DoEvents();
        }
    }
}
