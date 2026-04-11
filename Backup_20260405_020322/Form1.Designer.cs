namespace ExcelSQLiteWeb;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
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

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();
        webView21 = new Microsoft.Web.WebView2.WinForms.WebView2();
        ((System.ComponentModel.ISupportInitialize)webView21).BeginInit();
        SuspendLayout();
        // 
        // webView21
        // 
        webView21.AllowExternalDrop = true;
        webView21.CreationProperties = null;
        webView21.DefaultBackgroundColor = System.Drawing.Color.White;
        webView21.Dock = System.Windows.Forms.DockStyle.Fill;
        webView21.Location = new System.Drawing.Point(0, 0);
        webView21.Name = "webView21";
        webView21.Size = new System.Drawing.Size(1024, 768);
        webView21.Source = new System.Uri("file:///D:/A21-Trae项目/ExcelSQLite/ExcelSQLiteWeb/index.html", System.UriKind.Absolute);
        webView21.TabIndex = 0;
        webView21.ZoomFactor = 1D;
        // 
        // Form1
        // 
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new System.Drawing.Size(1024, 768);
        Controls.Add(webView21);
        Name = "ExcelSQLite";
        Text = "ExcelSQLite - 数据处理工具";
        WindowState = System.Windows.Forms.FormWindowState.Maximized;
        ((System.ComponentModel.ISupportInitialize)webView21).EndInit();
        ResumeLayout(false);
    }

    #endregion

    private Microsoft.Web.WebView2.WinForms.WebView2 webView21;
}
