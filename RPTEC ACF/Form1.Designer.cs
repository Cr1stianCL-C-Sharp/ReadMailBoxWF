using System;
using System.Windows.Forms;

namespace RPTEC_ACF
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);


            bool cursorNotInBar = Screen.GetWorkingArea(this).Contains(Cursor.Position);

            if (this.WindowState== FormWindowState.Minimized && cursorNotInBar)
            {
                NotifyIcon nt = new NotifyIcon();

                this.ShowInTaskbar = false;
                nt.Visible = true;
                this.Hide();

            }


        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.lbEstado = new System.Windows.Forms.Label();
            this.lbProcessState = new System.Windows.Forms.Label();
            this.txtTimer = new System.Windows.Forms.TextBox();
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.chkboxCiclo = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(19, 18);
            this.button1.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(284, 126);
            this.button1.TabIndex = 0;
            this.button1.Text = "INICIAR DEPURADOR";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.button2.Location = new System.Drawing.Point(19, 259);
            this.button2.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(284, 42);
            this.button2.TabIndex = 1;
            this.button2.Text = "DETENER";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // lbEstado
            // 
            this.lbEstado.AutoSize = true;
            this.lbEstado.Location = new System.Drawing.Point(12, 170);
            this.lbEstado.Name = "lbEstado";
            this.lbEstado.Size = new System.Drawing.Size(162, 39);
            this.lbEstado.TabIndex = 2;
            this.lbEstado.Text = "Estado  :";
            // 
            // lbProcessState
            // 
            this.lbProcessState.AutoSize = true;
            this.lbProcessState.ForeColor = System.Drawing.Color.Red;
            this.lbProcessState.Location = new System.Drawing.Point(198, 170);
            this.lbProcessState.Name = "lbProcessState";
            this.lbProcessState.Size = new System.Drawing.Size(200, 39);
            this.lbProcessState.TabIndex = 3;
            this.lbProcessState.Text = "No Iniciado";
            // 
            // txtTimer
            // 
            this.txtTimer.Location = new System.Drawing.Point(59, 350);
            this.txtTimer.Name = "txtTimer";
            this.txtTimer.Size = new System.Drawing.Size(450, 47);
            this.txtTimer.TabIndex = 4;
            // 
            // timer
            // 
            this.timer.Tick += new System.EventHandler(this.timer_Tick);
            // 
            // chkboxCiclo
            // 
            this.chkboxCiclo.AutoSize = true;
            this.chkboxCiclo.Enabled = false;
            this.chkboxCiclo.Location = new System.Drawing.Point(380, 61);
            this.chkboxCiclo.Name = "chkboxCiclo";
            this.chkboxCiclo.Size = new System.Drawing.Size(119, 43);
            this.chkboxCiclo.TabIndex = 5;
            this.chkboxCiclo.Text = "Ciclo";
            this.chkboxCiclo.UseVisualStyleBackColor = true;
            this.chkboxCiclo.CheckedChanged += new System.EventHandler(this.chkboxCiclo_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(21F, 39F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(655, 409);
            this.Controls.Add(this.chkboxCiclo);
            this.Controls.Add(this.txtTimer);
            this.Controls.Add(this.lbProcessState);
            this.Controls.Add(this.lbEstado);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 26.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.Name = "Form1";
            this.Text = "Mail Depurador Acf";
            this.Load += new System.EventHandler(this.Form1_Load_1);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label lbEstado;
        private System.Windows.Forms.Label lbProcessState;
        private TextBox txtTimer;
        private Timer timer;
        private CheckBox chkboxCiclo;
    }
}

