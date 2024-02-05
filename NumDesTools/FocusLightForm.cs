using System;
using System.Drawing;
using System.Windows.Forms;
using IWin32Window = System.Windows.Forms.IWin32Window;

namespace NumDesTools
{
    public  class FocusLightForm : Form
    {
        Win32Window _owner;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        private void InitializeComponent()
        {
            SuspendLayout();
            // 
            // FocusLightForm
            // 
            AutoScaleMode = AutoScaleMode.Inherit;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            BackColor = SystemColors.ActiveBorder;
            //透明度
            Opacity = 1;
            ClientSize = new Size(282, 153);
            ControlBox = false;
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            //边框粗细
            Padding = new Padding(2);
            Name = "FocusLightForm";
            ShowIcon = false;
            ShowInTaskbar = false;
            Load += FocusLightForm_Load;
            ResumeLayout(false);
        }

        public FocusLightForm()
        {
            InitializeComponent();
        }

        public void ShowToolTip()
        {
            //获取工作区句柄,显示窗口，不使用次方法，窗口闪现
            IntPtr excelHandle = (IntPtr)NumDesAddIn.App.Hwnd;
            _owner = new Win32Window(excelHandle);
            Show(_owner);
        }

        private void FocusLightForm_Load(object sender, EventArgs e)
        {

        }
        class Win32Window : IWin32Window
        {
            public IntPtr Handle
            {
                get;
                private set;
            }
            public Win32Window(IntPtr handle)
            {
                Handle = handle;
            }
        }
    }
}
