namespace Chamomile
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.тОЭToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.раздел1ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.раздел2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.раздел3ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.диагностикаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.экономикаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contentPanel = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.menuStrip1.SuspendLayout();
            this.contentPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.menuStrip1.Font = new System.Drawing.Font("Segoe UI", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.menuStrip1.GripMargin = new System.Windows.Forms.Padding(2, 2, 0, 2);
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.тОЭToolStripMenuItem,
            this.диагностикаToolStripMenuItem,
            this.экономикаToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(22, 7, 0, 7);
            this.menuStrip1.Size = new System.Drawing.Size(971, 52);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // тОЭToolStripMenuItem
            // 
            this.тОЭToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.раздел1ToolStripMenuItem,
            this.раздел2ToolStripMenuItem,
            this.раздел3ToolStripMenuItem});
            this.тОЭToolStripMenuItem.Name = "тОЭToolStripMenuItem";
            this.тОЭToolStripMenuItem.Size = new System.Drawing.Size(73, 38);
            this.тОЭToolStripMenuItem.Text = "ТОЭ";
            // 
            // раздел1ToolStripMenuItem
            // 
            this.раздел1ToolStripMenuItem.Name = "раздел1ToolStripMenuItem";
            this.раздел1ToolStripMenuItem.Size = new System.Drawing.Size(227, 44);
            this.раздел1ToolStripMenuItem.Text = "Раздел 1";
            this.раздел1ToolStripMenuItem.Click += new System.EventHandler(this.раздел1ToolStripMenuItem_Click);
            // 
            // раздел2ToolStripMenuItem
            // 
            this.раздел2ToolStripMenuItem.Name = "раздел2ToolStripMenuItem";
            this.раздел2ToolStripMenuItem.Size = new System.Drawing.Size(227, 44);
            this.раздел2ToolStripMenuItem.Text = "Раздел 2";
            this.раздел2ToolStripMenuItem.Click += new System.EventHandler(this.раздел2ToolStripMenuItem_Click);
            // 
            // раздел3ToolStripMenuItem
            // 
            this.раздел3ToolStripMenuItem.Name = "раздел3ToolStripMenuItem";
            this.раздел3ToolStripMenuItem.Size = new System.Drawing.Size(227, 44);
            this.раздел3ToolStripMenuItem.Text = "Раздел 3";
            this.раздел3ToolStripMenuItem.Click += new System.EventHandler(this.раздел3ToolStripMenuItem_Click);
            // 
            // диагностикаToolStripMenuItem
            // 
            this.диагностикаToolStripMenuItem.Name = "диагностикаToolStripMenuItem";
            this.диагностикаToolStripMenuItem.Size = new System.Drawing.Size(155, 38);
            this.диагностикаToolStripMenuItem.Text = "Диагностика";
            this.диагностикаToolStripMenuItem.Click += new System.EventHandler(this.диагностикаToolStripMenuItem_Click);
            // 
            // экономикаToolStripMenuItem
            // 
            this.экономикаToolStripMenuItem.Name = "экономикаToolStripMenuItem";
            this.экономикаToolStripMenuItem.Size = new System.Drawing.Size(219, 38);
            this.экономикаToolStripMenuItem.Text = "Экономика диплом";
            this.экономикаToolStripMenuItem.Click += new System.EventHandler(this.экономикаToolStripMenuItem_Click);
            // 
            // contentPanel
            // 
            this.contentPanel.Controls.Add(this.pictureBox1);
            this.contentPanel.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.contentPanel.Location = new System.Drawing.Point(15, 36);
            this.contentPanel.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.contentPanel.Name = "contentPanel";
            this.contentPanel.Size = new System.Drawing.Size(942, 512);
            this.contentPanel.TabIndex = 1;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(232, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(516, 545);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(22F, 42F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkGray;
            this.ClientSize = new System.Drawing.Size(971, 554);
            this.Controls.Add(this.contentPanel);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "Chamomile";
            this.TransparencyKey = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.contentPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem тОЭToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem диагностикаToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem экономикаToolStripMenuItem;
        private System.Windows.Forms.Panel contentPanel;
        private System.Windows.Forms.ToolStripMenuItem раздел1ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem раздел2ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem раздел3ToolStripMenuItem;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

