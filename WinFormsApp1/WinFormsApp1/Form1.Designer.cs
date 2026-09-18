namespace WinFormsApp1
{
    partial class lblBaseDatos
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
            lblbdcon = new Label();
            lblUsuario = new Label();
            tabControl1 = new TabControl();
            tabPage1 = new TabPage();
            lblProgreso = new Label();
            progressBar1 = new ProgressBar();
            label2 = new Label();
            button6 = new Button();
            button4 = new Button();
            dataGridView2 = new DataGridView();
            label1 = new Label();
            button5 = new Button();
            button3 = new Button();
            dataGridView1 = new DataGridView();
            button9 = new Button();
            button2 = new Button();
            dgv3 = new DataGridView();
            progressBar3 = new ProgressBar();
            progressBar2 = new ProgressBar();
            button10 = new Button();
            button8 = new Button();
            button7 = new Button();
            tabPage2 = new TabPage();
            label3 = new Label();
            progressBarFase2 = new ProgressBar();
            button13 = new Button();
            button12 = new Button();
            button11 = new Button();
            dgvFase2 = new DataGridView();
            tabPage3 = new TabPage();
            label4 = new Label();
            progressBarF3 = new ProgressBar();
            crear_ubicacion = new Button();
            vali_crear_subniveles = new Button();
            importExcel = new Button();
            consultarUbiBD_F3 = new Button();
            dgv_y = new DataGridView();
            dgv_x = new DataGridView();
            tabPage4 = new TabPage();
            f4_importarExcel_2 = new Button();
            f4_actualizarmasiva_2 = new Button();
            f4_prevalidacion_2 = new Button();
            f4_exportarlog = new Button();
            label5 = new Label();
            f4_progressBar = new ProgressBar();
            f4_actualizarmasiva = new Button();
            f4_prevalidacion = new Button();
            f4_importarExcel = new Button();
            f4_dgv = new DataGridView();
            textBox1 = new TextBox();
            button1 = new Button();
            label6 = new Label();
            label7 = new Label();
            tabControl1.SuspendLayout();
            tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgv3).BeginInit();
            tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvFase2).BeginInit();
            tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgv_y).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgv_x).BeginInit();
            tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)f4_dgv).BeginInit();
            SuspendLayout();
            // 
            // lblbdcon
            // 
            lblbdcon.AutoSize = true;
            lblbdcon.Location = new Point(724, 41);
            lblbdcon.Name = "lblbdcon";
            lblbdcon.Size = new Size(12, 15);
            lblbdcon.TabIndex = 11;
            lblbdcon.Text = "-";
            // 
            // lblUsuario
            // 
            lblUsuario.AutoSize = true;
            lblUsuario.Location = new Point(900, 41);
            lblUsuario.Name = "lblUsuario";
            lblUsuario.Size = new Size(12, 15);
            lblUsuario.TabIndex = 12;
            lblUsuario.Text = "-";
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Controls.Add(tabPage3);
            tabControl1.Controls.Add(tabPage4);
            tabControl1.Location = new Point(12, 78);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(1059, 674);
            tabControl1.TabIndex = 22;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(lblProgreso);
            tabPage1.Controls.Add(progressBar1);
            tabPage1.Controls.Add(label2);
            tabPage1.Controls.Add(button6);
            tabPage1.Controls.Add(button4);
            tabPage1.Controls.Add(dataGridView2);
            tabPage1.Controls.Add(label1);
            tabPage1.Controls.Add(button5);
            tabPage1.Controls.Add(button3);
            tabPage1.Controls.Add(dataGridView1);
            tabPage1.Controls.Add(button9);
            tabPage1.Controls.Add(button2);
            tabPage1.Controls.Add(dgv3);
            tabPage1.Controls.Add(progressBar3);
            tabPage1.Controls.Add(progressBar2);
            tabPage1.Controls.Add(button10);
            tabPage1.Controls.Add(button8);
            tabPage1.Controls.Add(button7);
            tabPage1.Location = new Point(4, 24);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(1051, 646);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "Fase1";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // lblProgreso
            // 
            lblProgreso.AutoSize = true;
            lblProgreso.Location = new Point(384, 645);
            lblProgreso.Name = "lblProgreso";
            lblProgreso.Size = new Size(22, 15);
            lblProgreso.TabIndex = 38;
            lblProgreso.Text = "---";
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(382, 617);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(387, 23);
            progressBar1.TabIndex = 37;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(244, 591);
            label2.Name = "label2";
            label2.Size = new Size(79, 15);
            label2.TabIndex = 36;
            label2.Text = "Se encontro : ";
            // 
            // button6
            // 
            button6.Location = new Point(10, 587);
            button6.Name = "button6";
            button6.Size = new Size(208, 23);
            button6.TabIndex = 35;
            button6.Text = "Sincronizar Subniveles Faltantes";
            button6.UseVisualStyleBackColor = true;
            button6.Click += button6_Click_1;
            // 
            // button4
            // 
            button4.Location = new Point(10, 617);
            button4.Name = "button4";
            button4.Size = new Size(330, 23);
            button4.TabIndex = 34;
            button4.Text = "Actualizar/Crear en BD";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click_1;
            // 
            // dataGridView2
            // 
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView2.Location = new Point(11, 462);
            dataGridView2.Name = "dataGridView2";
            dataGridView2.Size = new Size(759, 109);
            dataGridView2.TabIndex = 33;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(153, 409);
            label1.Name = "label1";
            label1.Size = new Size(79, 15);
            label1.TabIndex = 32;
            label1.Text = "Se encontro : ";
            // 
            // button5
            // 
            button5.Location = new Point(10, 405);
            button5.Name = "button5";
            button5.Size = new Size(121, 23);
            button5.TabIndex = 31;
            button5.Text = "Limpiar_dgv1";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // button3
            // 
            button3.Location = new Point(11, 433);
            button3.Name = "button3";
            button3.Size = new Size(330, 23);
            button3.TabIndex = 30;
            button3.Text = "Importar Ubicaciones Finales/Reales";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click_1;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(8, 251);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Size = new Size(635, 141);
            dataGridView1.TabIndex = 29;
            // 
            // button9
            // 
            button9.Location = new Point(8, 222);
            button9.Name = "button9";
            button9.Size = new Size(238, 23);
            button9.TabIndex = 28;
            button9.Text = "Limpiar ProgressBar";
            button9.UseVisualStyleBackColor = true;
            button9.Click += button9_Click_1;
            // 
            // button2
            // 
            button2.Location = new Point(378, 222);
            button2.Name = "button2";
            button2.Size = new Size(330, 23);
            button2.TabIndex = 27;
            button2.Text = "Consulta Ubicaciones en BD";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click_1;
            // 
            // dgv3
            // 
            dgv3.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgv3.Location = new Point(8, 66);
            dgv3.Name = "dgv3";
            dgv3.Size = new Size(635, 150);
            dgv3.TabIndex = 26;
            // 
            // progressBar3
            // 
            progressBar3.Location = new Point(680, 169);
            progressBar3.Name = "progressBar3";
            progressBar3.Size = new Size(351, 23);
            progressBar3.TabIndex = 25;
            // 
            // progressBar2
            // 
            progressBar2.Location = new Point(680, 87);
            progressBar2.Name = "progressBar2";
            progressBar2.Size = new Size(351, 23);
            progressBar2.TabIndex = 24;
            // 
            // button10
            // 
            button10.Location = new Point(680, 48);
            button10.Name = "button10";
            button10.Size = new Size(216, 23);
            button10.TabIndex = 23;
            button10.Text = "Activar Articulos";
            button10.UseVisualStyleBackColor = true;
            button10.Click += button10_Click_1;
            // 
            // button8
            // 
            button8.Location = new Point(680, 127);
            button8.Name = "button8";
            button8.Size = new Size(202, 23);
            button8.TabIndex = 22;
            button8.Text = "Reubicacion Stock";
            button8.UseVisualStyleBackColor = true;
            button8.Click += button8_Click_1;
            // 
            // button7
            // 
            button7.Location = new Point(6, 37);
            button7.Name = "button7";
            button7.Size = new Size(172, 23);
            button7.TabIndex = 17;
            button7.Text = "Consular UbicacionStock";
            button7.UseVisualStyleBackColor = true;
            button7.Click += button7_Click_1;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(label3);
            tabPage2.Controls.Add(progressBarFase2);
            tabPage2.Controls.Add(button13);
            tabPage2.Controls.Add(button12);
            tabPage2.Controls.Add(button11);
            tabPage2.Controls.Add(dgvFase2);
            tabPage2.Location = new Point(4, 24);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(1051, 646);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "Fase2";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(309, 271);
            label3.Name = "label3";
            label3.Size = new Size(38, 15);
            label3.TabIndex = 5;
            label3.Text = "label3";
            // 
            // progressBarFase2
            // 
            progressBarFase2.Location = new Point(15, 267);
            progressBarFase2.Name = "progressBarFase2";
            progressBarFase2.Size = new Size(271, 23);
            progressBarFase2.TabIndex = 4;
            // 
            // button13
            // 
            button13.Location = new Point(809, 120);
            button13.Name = "button13";
            button13.Size = new Size(192, 23);
            button13.TabIndex = 3;
            button13.Text = "3. Ejecutar Putaway (Reubicar)";
            button13.UseVisualStyleBackColor = true;
            button13.Click += button13_Click;
            // 
            // button12
            // 
            button12.Location = new Point(809, 75);
            button12.Name = "button12";
            button12.Size = new Size(192, 23);
            button12.TabIndex = 2;
            button12.Text = "2. Pre-Validar Distribución";
            button12.UseVisualStyleBackColor = true;
            button12.Click += button12_Click;
            // 
            // button11
            // 
            button11.Location = new Point(809, 28);
            button11.Name = "button11";
            button11.Size = new Size(192, 23);
            button11.TabIndex = 1;
            button11.Text = "1. Importar Layout Picking";
            button11.UseVisualStyleBackColor = true;
            button11.Click += button11_Click;
            // 
            // dgvFase2
            // 
            dgvFase2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvFase2.Location = new Point(15, 28);
            dgvFase2.Name = "dgvFase2";
            dgvFase2.Size = new Size(750, 220);
            dgvFase2.TabIndex = 0;
            // 
            // tabPage3
            // 
            tabPage3.Controls.Add(label4);
            tabPage3.Controls.Add(progressBarF3);
            tabPage3.Controls.Add(crear_ubicacion);
            tabPage3.Controls.Add(vali_crear_subniveles);
            tabPage3.Controls.Add(importExcel);
            tabPage3.Controls.Add(consultarUbiBD_F3);
            tabPage3.Controls.Add(dgv_y);
            tabPage3.Controls.Add(dgv_x);
            tabPage3.Location = new Point(4, 24);
            tabPage3.Name = "tabPage3";
            tabPage3.Padding = new Padding(3);
            tabPage3.Size = new Size(1051, 646);
            tabPage3.TabIndex = 2;
            tabPage3.Text = "Fase3";
            tabPage3.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(23, 365);
            label4.Name = "label4";
            label4.Size = new Size(38, 15);
            label4.TabIndex = 39;
            label4.Text = "label4";
            // 
            // progressBarF3
            // 
            progressBarF3.Location = new Point(20, 388);
            progressBarF3.Name = "progressBarF3";
            progressBarF3.Size = new Size(622, 23);
            progressBarF3.TabIndex = 38;
            // 
            // crear_ubicacion
            // 
            crear_ubicacion.Location = new Point(531, 336);
            crear_ubicacion.Name = "crear_ubicacion";
            crear_ubicacion.Size = new Size(111, 23);
            crear_ubicacion.TabIndex = 5;
            crear_ubicacion.Text = "Crear Ubicaciones";
            crear_ubicacion.UseVisualStyleBackColor = true;
            crear_ubicacion.Click += crear_ubicacion_Click;
            // 
            // vali_crear_subniveles
            // 
            vali_crear_subniveles.Location = new Point(20, 336);
            vali_crear_subniveles.Name = "vali_crear_subniveles";
            vali_crear_subniveles.Size = new Size(173, 23);
            vali_crear_subniveles.TabIndex = 4;
            vali_crear_subniveles.Text = "Validar crear Subniveles";
            vali_crear_subniveles.UseVisualStyleBackColor = true;
            vali_crear_subniveles.Click += vali_crear_subniveles_Click;
            // 
            // importExcel
            // 
            importExcel.Location = new Point(20, 171);
            importExcel.Name = "importExcel";
            importExcel.Size = new Size(284, 23);
            importExcel.TabIndex = 3;
            importExcel.Text = "Importar/Leer Excel Ubicacion por Crear";
            importExcel.UseVisualStyleBackColor = true;
            importExcel.Click += importExcel_Click;
            // 
            // consultarUbiBD_F3
            // 
            consultarUbiBD_F3.Location = new Point(20, 7);
            consultarUbiBD_F3.Name = "consultarUbiBD_F3";
            consultarUbiBD_F3.Size = new Size(149, 23);
            consultarUbiBD_F3.TabIndex = 2;
            consultarUbiBD_F3.Text = "ConsultarUbicacion BD";
            consultarUbiBD_F3.UseVisualStyleBackColor = true;
            consultarUbiBD_F3.Click += consultarUbiBD_F3_Click;
            // 
            // dgv_y
            // 
            dgv_y.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgv_y.Location = new Point(20, 200);
            dgv_y.Name = "dgv_y";
            dgv_y.Size = new Size(622, 130);
            dgv_y.TabIndex = 1;
            // 
            // dgv_x
            // 
            dgv_x.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgv_x.Location = new Point(20, 36);
            dgv_x.Name = "dgv_x";
            dgv_x.Size = new Size(622, 129);
            dgv_x.TabIndex = 0;
            // 
            // tabPage4
            // 
            tabPage4.Controls.Add(f4_importarExcel_2);
            tabPage4.Controls.Add(f4_actualizarmasiva_2);
            tabPage4.Controls.Add(f4_prevalidacion_2);
            tabPage4.Controls.Add(f4_exportarlog);
            tabPage4.Controls.Add(label5);
            tabPage4.Controls.Add(f4_progressBar);
            tabPage4.Controls.Add(f4_actualizarmasiva);
            tabPage4.Controls.Add(f4_prevalidacion);
            tabPage4.Controls.Add(f4_importarExcel);
            tabPage4.Controls.Add(f4_dgv);
            tabPage4.Location = new Point(4, 24);
            tabPage4.Name = "tabPage4";
            tabPage4.Size = new Size(1051, 646);
            tabPage4.TabIndex = 3;
            tabPage4.Text = "Fase4";
            tabPage4.UseVisualStyleBackColor = true;
            // 
            // f4_importarExcel_2
            // 
            f4_importarExcel_2.Location = new Point(567, 14);
            f4_importarExcel_2.Name = "f4_importarExcel_2";
            f4_importarExcel_2.Size = new Size(267, 23);
            f4_importarExcel_2.TabIndex = 9;
            f4_importarExcel_2.Text = "Importar Excel_2";
            f4_importarExcel_2.UseVisualStyleBackColor = true;
            f4_importarExcel_2.Click += button14_Click;
            // 
            // f4_actualizarmasiva_2
            // 
            f4_actualizarmasiva_2.Location = new Point(627, 325);
            f4_actualizarmasiva_2.Name = "f4_actualizarmasiva_2";
            f4_actualizarmasiva_2.Size = new Size(207, 23);
            f4_actualizarmasiva_2.TabIndex = 8;
            f4_actualizarmasiva_2.Text = "actualizacion_Masiva_2";
            f4_actualizarmasiva_2.UseVisualStyleBackColor = true;
            f4_actualizarmasiva_2.Click += f4_actualizarmasiva_2_Click;
            // 
            // f4_prevalidacion_2
            // 
            f4_prevalidacion_2.Location = new Point(17, 325);
            f4_prevalidacion_2.Name = "f4_prevalidacion_2";
            f4_prevalidacion_2.Size = new Size(207, 23);
            f4_prevalidacion_2.TabIndex = 7;
            f4_prevalidacion_2.Text = "pre-validacion_2";
            f4_prevalidacion_2.UseVisualStyleBackColor = true;
            f4_prevalidacion_2.Click += f4_prevalidacion_2_Click;
            // 
            // f4_exportarlog
            // 
            f4_exportarlog.Location = new Point(354, 275);
            f4_exportarlog.Name = "f4_exportarlog";
            f4_exportarlog.Size = new Size(127, 23);
            f4_exportarlog.TabIndex = 6;
            f4_exportarlog.Text = "Exportar_log";
            f4_exportarlog.UseVisualStyleBackColor = true;
            f4_exportarlog.Click += f4_exportarlog_Click;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(25, 307);
            label5.Name = "label5";
            label5.Size = new Size(38, 15);
            label5.TabIndex = 5;
            label5.Text = "label5";
            // 
            // f4_progressBar
            // 
            f4_progressBar.Location = new Point(167, 369);
            f4_progressBar.Name = "f4_progressBar";
            f4_progressBar.Size = new Size(667, 23);
            f4_progressBar.TabIndex = 4;
            // 
            // f4_actualizarmasiva
            // 
            f4_actualizarmasiva.Location = new Point(627, 275);
            f4_actualizarmasiva.Name = "f4_actualizarmasiva";
            f4_actualizarmasiva.Size = new Size(207, 23);
            f4_actualizarmasiva.TabIndex = 3;
            f4_actualizarmasiva.Text = "actualizacion_Masiva";
            f4_actualizarmasiva.UseVisualStyleBackColor = true;
            f4_actualizarmasiva.Click += f4_actualizarmasiva_Click;
            // 
            // f4_prevalidacion
            // 
            f4_prevalidacion.Location = new Point(17, 275);
            f4_prevalidacion.Name = "f4_prevalidacion";
            f4_prevalidacion.Size = new Size(207, 23);
            f4_prevalidacion.TabIndex = 2;
            f4_prevalidacion.Text = "pre-validacion";
            f4_prevalidacion.UseVisualStyleBackColor = true;
            f4_prevalidacion.Click += f4_prevalidacion_Click;
            // 
            // f4_importarExcel
            // 
            f4_importarExcel.Location = new Point(17, 14);
            f4_importarExcel.Name = "f4_importarExcel";
            f4_importarExcel.Size = new Size(267, 23);
            f4_importarExcel.TabIndex = 1;
            f4_importarExcel.Text = "Importar Excel";
            f4_importarExcel.UseVisualStyleBackColor = true;
            f4_importarExcel.Click += f4_importarExcel_Click;
            // 
            // f4_dgv
            // 
            f4_dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            f4_dgv.Location = new Point(17, 43);
            f4_dgv.Name = "f4_dgv";
            f4_dgv.Size = new Size(816, 226);
            f4_dgv.TabIndex = 0;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(16, 6);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(692, 23);
            textBox1.TabIndex = 24;
            // 
            // button1
            // 
            button1.Location = new Point(752, 3);
            button1.Name = "button1";
            button1.Size = new Size(112, 23);
            button1.TabIndex = 23;
            button1.Text = "Generar Tocken";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click_2;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(269, 41);
            label6.Name = "label6";
            label6.Size = new Size(38, 15);
            label6.TabIndex = 25;
            label6.Text = "label6";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(528, 41);
            label7.Name = "label7";
            label7.Size = new Size(38, 15);
            label7.TabIndex = 26;
            label7.Text = "label7";
            // 
            // lblBaseDatos
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoScroll = true;
            ClientSize = new Size(1284, 749);
            Controls.Add(label7);
            Controls.Add(label6);
            Controls.Add(textBox1);
            Controls.Add(button1);
            Controls.Add(tabControl1);
            Controls.Add(lblUsuario);
            Controls.Add(lblbdcon);
            Name = "lblBaseDatos";
            Text = "Form1";
            Load += Form1_Load;
            tabControl1.ResumeLayout(false);
            tabPage1.ResumeLayout(false);
            tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgv3).EndInit();
            tabPage2.ResumeLayout(false);
            tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgvFase2).EndInit();
            tabPage3.ResumeLayout(false);
            tabPage3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgv_y).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgv_x).EndInit();
            tabPage4.ResumeLayout(false);
            tabPage4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)f4_dgv).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Label lblbdcon;
        private Label lblUsuario;
        private TabControl tabControl1;
        private TabPage tabPage1;
        private TabPage tabPage2;
        private Button button7;
        private Label label1;
        private Button button5;
        private Button button3;
        private DataGridView dataGridView1;
        private Button button9;
        private Button button2;
        private DataGridView dgv3;
        private ProgressBar progressBar3;
        private ProgressBar progressBar2;
        private Button button10;
        private Button button8;
        private Label lblProgreso;
        private ProgressBar progressBar1;
        private Label label2;
        private Button button6;
        private Button button4;
        private DataGridView dataGridView2;
        private TabPage tabPage3;
        private TextBox textBox1;
        private Button button1;
        private Button button13;
        private Button button12;
        private Button button11;
        private DataGridView dgvFase2;
        private Label label3;
        private ProgressBar progressBarFase2;
        private Button consultarUbiBD_F3;
        private DataGridView dgv_y;
        private DataGridView dgv_x;
        private Button importExcel;
        private Button crear_ubicacion;
        private Button vali_crear_subniveles;
        private Label label4;
        private ProgressBar progressBarF3;
        private TabPage tabPage4;
        private Button f4_importarExcel;
        private DataGridView f4_dgv;
        private ProgressBar f4_progressBar;
        private Button f4_actualizarmasiva;
        private Button f4_prevalidacion;
        private Button f4_exportarlog;
        private Label label5;
        private Label label6;
        private Label label7;
        private Button f4_actualizarmasiva_2;
        private Button f4_prevalidacion_2;
        private Button f4_importarExcel_2;
    }
}
