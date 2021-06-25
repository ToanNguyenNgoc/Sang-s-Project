namespace TVMH
{
    partial class frmMain
    {
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.tưVấnNângCaoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dùngMạngNơronToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.tưVấnNângCaoToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(907, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(12, 20);
            // 
            // tưVấnNângCaoToolStripMenuItem
            // 
            this.tưVấnNângCaoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.dùngMạngNơronToolStripMenuItem});
            this.tưVấnNângCaoToolStripMenuItem.Name = "tưVấnNângCaoToolStripMenuItem";
            this.tưVấnNângCaoToolStripMenuItem.Size = new System.Drawing.Size(107, 20);
            this.tưVấnNângCaoToolStripMenuItem.Text = "Tư vấn nâng cao";
            // 
            // dùngMạngNơronToolStripMenuItem
            // 
            this.dùngMạngNơronToolStripMenuItem.Name = "dùngMạngNơronToolStripMenuItem";
            this.dùngMạngNơronToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.dùngMạngNơronToolStripMenuItem.Text = "Tư vấn môn học";
            this.dùngMạngNơronToolStripMenuItem.Click += new System.EventHandler(this.dùngMạngNơronToolStripMenuItem_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.ClientSize = new System.Drawing.Size(907, 420);
            this.Controls.Add(this.menuStrip1);
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmMain";
            this.Text = "TƯ VẤN HỌC TẬP CHO SINH VIÊN TRƯỜNG ĐẠI HỌC ";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem tưVấnNângCaoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dùngMạngNơronToolStripMenuItem;
    }
}