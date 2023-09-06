using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Deployment.Application;

namespace KMIQ
{
    partial class AboutBox1 : System.Windows.Forms.Form
    {
        public AboutBox1()
        {
            InitializeComponent();
            this.Text = String.Format("{0} 정보", AssemblyTitle);
            this.labelProductName.Text = string.Format("제품 이름 : {0}", AssemblyProduct);
            this.labelVersion.Text = String.Format("제품 버전 : {0}", AssemblyVersion);
            this.labelCopyright.Text = string.Format("저작권 : {0}",AssemblyCopyright);
            this.labelCompanyName.Text = string.Format("제작 회사 : {0}",AssemblyCompany);
            this.textBoxDescription.Text = string.Format("{0}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}{1}실행경로 : {2}{1}{1}{1}Data경로 : {3}", AssemblyDescription, Environment.NewLine, AppDomain.CurrentDomain.BaseDirectory, (ApplicationDeployment.IsNetworkDeployed ? ApplicationDeployment.CurrentDeployment.DataDirectory.ToString() : AppDomain.CurrentDomain.BaseDirectory));
        }

        #region 어셈블리 특성 접근자

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion

        private void OKButton_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel.Link link = new LinkLabel.Link();
            link.LinkData = "http://www.smartomr.co.kr";
            LinkLabel1.Links.Add(link);

            System.Diagnostics.Process.Start(link.LinkData.ToString());
        }
    }
}
