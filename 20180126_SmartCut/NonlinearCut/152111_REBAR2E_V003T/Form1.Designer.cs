namespace REBAR2ET
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this._TXTETABS = new System.Windows.Forms.TextBox();
            this._COL_REBAR = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this._BEAM_REBAR = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this._COL_hinge_i = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this._COL_hinge_j = new System.Windows.Forms.TextBox();
            this._BEAM_hinge_i = new System.Windows.Forms.TextBox();
            this._BEAM_hinge_j = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this._COL_COVER = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this._BEAM_COVER = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(602, 188);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "開始";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(631, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "Ver.0.03";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // _TXTETABS
            // 
            this._TXTETABS.Location = new System.Drawing.Point(12, 38);
            this._TXTETABS.Name = "_TXTETABS";
            this._TXTETABS.Size = new System.Drawing.Size(192, 22);
            this._TXTETABS.TabIndex = 2;
            // 
            // _COL_REBAR
            // 
            this._COL_REBAR.Location = new System.Drawing.Point(12, 91);
            this._COL_REBAR.Name = "_COL_REBAR";
            this._COL_REBAR.Size = new System.Drawing.Size(192, 22);
            this._COL_REBAR.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(133, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "ETABS文字檔(UNIT:t ,m)";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 12);
            this.label3.TabIndex = 5;
            this.label3.Text = "COL_REBAR(*.col)";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(210, 91);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 6;
            this.button2.Text = "選擇檔案";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(210, 39);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 7;
            this.button3.Text = "選擇檔案";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(212, 144);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 23);
            this.button4.TabIndex = 10;
            this.button4.Text = "選擇檔案";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(14, 126);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(125, 12);
            this.label4.TabIndex = 9;
            this.label4.Text = "BEAM_REBAR(*.beam)";
            // 
            // _BEAM_REBAR
            // 
            this._BEAM_REBAR.Location = new System.Drawing.Point(14, 144);
            this._BEAM_REBAR.Name = "_BEAM_REBAR";
            this._BEAM_REBAR.Size = new System.Drawing.Size(190, 22);
            this._BEAM_REBAR.TabIndex = 8;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 35);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(118, 12);
            this.label5.TabIndex = 11;
            this.label5.Text = "柱 i 端塑鉸位置(相對)";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(195, 35);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(118, 12);
            this.label6.TabIndex = 12;
            this.label6.Text = "梁 i 端塑鉸位置(相對)";
            // 
            // _COL_hinge_i
            // 
            this._COL_hinge_i.Location = new System.Drawing.Point(131, 28);
            this._COL_hinge_i.Name = "_COL_hinge_i";
            this._COL_hinge_i.Size = new System.Drawing.Size(44, 22);
            this._COL_hinge_i.TabIndex = 13;
            this._COL_hinge_i.Text = "0.1";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 73);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(118, 12);
            this.label7.TabIndex = 14;
            this.label7.Text = "柱 j 端塑鉸位置(相對)";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(196, 73);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(118, 12);
            this.label8.TabIndex = 15;
            this.label8.Text = "梁 j 端塑鉸位置(相對)";
            // 
            // _COL_hinge_j
            // 
            this._COL_hinge_j.Location = new System.Drawing.Point(131, 67);
            this._COL_hinge_j.Name = "_COL_hinge_j";
            this._COL_hinge_j.Size = new System.Drawing.Size(44, 22);
            this._COL_hinge_j.TabIndex = 16;
            this._COL_hinge_j.Text = "0.8";
            // 
            // _BEAM_hinge_i
            // 
            this._BEAM_hinge_i.Location = new System.Drawing.Point(326, 26);
            this._BEAM_hinge_i.Name = "_BEAM_hinge_i";
            this._BEAM_hinge_i.Size = new System.Drawing.Size(44, 22);
            this._BEAM_hinge_i.TabIndex = 17;
            this._BEAM_hinge_i.Text = "0.1";
            // 
            // _BEAM_hinge_j
            // 
            this._BEAM_hinge_j.Location = new System.Drawing.Point(326, 64);
            this._BEAM_hinge_j.Name = "_BEAM_hinge_j";
            this._BEAM_hinge_j.Size = new System.Drawing.Size(44, 22);
            this._BEAM_hinge_j.TabIndex = 18;
            this._BEAM_hinge_j.Text = "0.9";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 111);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(94, 12);
            this.label9.TabIndex = 19;
            this.label9.Text = "柱保護層厚度(m)";
            // 
            // _COL_COVER
            // 
            this._COL_COVER.Location = new System.Drawing.Point(131, 108);
            this._COL_COVER.Name = "_COL_COVER";
            this._COL_COVER.Size = new System.Drawing.Size(44, 22);
            this._COL_COVER.TabIndex = 20;
            this._COL_COVER.Text = "0.07";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(201, 111);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(94, 12);
            this.label10.TabIndex = 21;
            this.label10.Text = "梁保護層厚度(m)";
            // 
            // _BEAM_COVER
            // 
            this._BEAM_COVER.Location = new System.Drawing.Point(617, 128);
            this._BEAM_COVER.Name = "_BEAM_COVER";
            this._BEAM_COVER.Size = new System.Drawing.Size(44, 22);
            this._BEAM_COVER.TabIndex = 22;
            this._BEAM_COVER.Text = "0.09";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this._BEAM_hinge_j);
            this.groupBox1.Controls.Add(this._COL_COVER);
            this.groupBox1.Controls.Add(this._BEAM_hinge_i);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this._COL_hinge_j);
            this.groupBox1.Controls.Add(this._COL_hinge_i);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Location = new System.Drawing.Point(291, 22);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(386, 151);
            this.groupBox1.TabIndex = 23;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "使用者輸入";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(682, 223);
            this.Controls.Add(this._BEAM_COVER);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label4);
            this.Controls.Add(this._BEAM_REBAR);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this._COL_REBAR);
            this.Controls.Add(this._TXTETABS);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "Form1";
            this.Text = "REBAR2ETABS";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox _TXTETABS;
        private System.Windows.Forms.TextBox _COL_REBAR;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox _BEAM_REBAR;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox _COL_hinge_i;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox _COL_hinge_j;
        private System.Windows.Forms.TextBox _BEAM_hinge_i;
        private System.Windows.Forms.TextBox _BEAM_hinge_j;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox _COL_COVER;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox _BEAM_COVER;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}

