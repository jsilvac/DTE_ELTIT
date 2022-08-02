using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SchoolManagementAdmin.objetos;
using MetroFramework;
using MySql.Data.MySqlClient;

namespace SchoolManagementAdmin
{
    public partial class frmListado : Form
    {
        public frmListado()
        {
            InitializeComponent();
        }

        public frmMain main;

        string[] month = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };

        //For animated panels direction
        string optionsDirection = "down";
        string toastDirection = "down";
        string rightDirection = "right";

        //For animated panels timeout
        int optionsTimeOut = 0;
        int toastTimeOut = 0;
        int RightTimeOut = 0;

        //For animated panels position
        int optionsX;
        int optionsY;
        int rightX;
        int rightY;

        //method to set fullscreen
        private void setFullScreen()
        {
            int x = Screen.PrimaryScreen.Bounds.Width;
            int y = Screen.PrimaryScreen.Bounds.Height - 10;
            Location = new Point(0, 0);
            Size = new Size(x, y);
        }

        //method to set the position of the main panel that holds the controls to center of the form.
        private void setMainPanelPosition()
        {
            int mX = (Width - pnlMain.Width) / 2;
            int mY = (Height - pnlMain.Height) / 2;
            pnlMain.Location = new Point(mX, mY);
        }

        //method to set the initial location of the options panel.


        private void setRightOptionsPanelPosition()
        {
            int y = Height;
            rightY = 0;
            rightX = Width + pnlRightOptions.Width;
            pnlRightOptions.Size = new Size(pnlRightOptions.Width, y);
            pnlRightOptions.Location = new Point(rightX, rightY);
            int rX = pnlRightMain.Location.X;
            int rY = (pnlRightOptions.Height - pnlRightMain.Height) / 2;
            pnlRightMain.Location = new Point(rX, rY);
        }





        private void frmTemplate_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Y >= Height - 15 && e.X < (Width - pnlRightOptions.Width))
            {
                optionsDirection = "up";
                rightDirection = "right";
                optionsTimeOut = 0;
            }
            if (e.X >= Width - 15)
            {
                rightDirection = "left";
                RightTimeOut = 0;
                optionsDirection = "down";
            }
            if (e.X < (Width - pnlRightOptions.Width))
            {
                rightDirection = "Left";
            }
        }



        private void RightOptions_Tick(object sender, EventArgs e)
        {
            if (RightTimeOut < 1000)
            {
                RightTimeOut++;
            }
            if (RightTimeOut == 1000)
            {
                if (rightDirection == "left")
                {
                    rightDirection = "right";
                }
            }
            if (rightDirection == "left")
            {
                if (rightX > Width - pnlRightOptions.Width)
                {
                    rightX -= 2;
                    pnlRightOptions.Location = new Point(rightX, rightY);
                }
            }
            else
            {
                if (rightX < Width)
                {
                    rightX += 2;
                }
                pnlRightOptions.Location = new Point(rightX, rightY);
            }
        }

        private void pbHide_Click(object sender, EventArgs e)
        {
            optionsDirection = "down";
        }

        private void pbExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pbHome_Click(object sender, EventArgs e)
        {
            main.Show();
            this.Close();
        }

        private void pnlMain_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnInicia_Click(object sender, EventArgs e)
        {


        }

        private void button3_Click(object sender, EventArgs e)
        {
            MySqlCommand cmd = null;
            Ventas v = new Ventas();
            cmd = v.GetNotasByVendedor(Inicial.G_LOCAL, Inicial.G_VENDEDOR, Inicial.G_CAJA, txtdia.Text, txtmes.Text, txtaño.Text);

            MySqlDataAdapter da = new MySqlDataAdapter(cmd);
            DataTable dt = new DataTable();

            listaNotas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            listaNotas.RowTemplate.Height = 130;
            listaNotas.AllowUserToAddRows = false;

            da.Fill(dt);

            listaNotas.DataSource = dt;
            da.Dispose();

            CargaGrillaNotas();
            lblmensaje.Text = "SE ENCONTRARON " + listaNotas.RowCount + " REGISTROS";
        }

        private void CargaGrillaNotas()
        {
            //productos.ColumnCount = 4;
            listaNotas.Columns[0].Width = 70;
            listaNotas.Columns[1].Width = 70;
            listaNotas.Columns[2].Width = 210;
            listaNotas.Columns[3].Width = 80;
            listaNotas.Columns[3].DefaultCellStyle.Format = "C0";

            foreach (DataGridViewRow row in listaNotas.Rows)
            {
                row.Height = 70;



                row.Cells[0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                row.Cells[1].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                row.Cells[2].Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                row.Cells[3].Style.Alignment = DataGridViewContentAlignment.MiddleRight;

                row.Cells[1].Style.Font = new Font(row.Cells[1].InheritedStyle.Font, FontStyle.Bold);
                row.Cells[2].Style.Font = new Font("Verdana", 8F, FontStyle.Bold);
                row.Cells[3].Style.Font = new Font(row.Cells[3].InheritedStyle.Font, FontStyle.Bold);
            }
            // productos.Columns[3].DefaultCellStyle.Format = "N2";
            // productos.Rows[0].Height = 0;
            listaNotas.ColumnHeadersVisible = false;
            listaNotas.Refresh();


        }

        private void txtmes_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = false;
            }
            else
             if (Char.IsControl(e.KeyChar)) //permitir teclas de control como retroceso 
            {
                e.Handled = false;
            }
            else
            {
                //el resto de teclas pulsadas se desactivan 
                e.Handled = true;
            }
        }

        private void txtaño_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = false;
            }
            else
             if (Char.IsControl(e.KeyChar)) //permitir teclas de control como retroceso 
            {
                e.Handled = false;
            }
            else
            {
                //el resto de teclas pulsadas se desactivan 
                e.Handled = true;
            }
        }

        private void txtmes_Leave(object sender, EventArgs e)
        {
            txtmes.Text = txtmes.Text.PadLeft(2, '0');
        }

        private void txtaño_Leave(object sender, EventArgs e)
        {
            txtaño.Text = txtaño.Text.PadLeft(4, '0');
        }

        private void frmListado_Load(object sender, EventArgs e)
        {
            txtdia.Text = DateTime.Now.ToString("dd");
            txtmes.Text = DateTime.Now.ToString("MM");
            txtaño.Text = DateTime.Now.ToString("yyyy");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
            //main.Show();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            MySqlCommand cmd = null;
            Ventas v = new Ventas();
            cmd = v.GetNotasByVendedor(Inicial.G_LOCAL, Inicial.G_VENDEDOR, Inicial.G_CAJA, this.txtdia.Text, txtmes.Text, txtaño.Text);

            MySqlDataAdapter da = new MySqlDataAdapter(cmd);
            DataTable dt = new DataTable();

            listaNotas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            listaNotas.RowTemplate.Height = 130;
            listaNotas.AllowUserToAddRows = false;

            da.Fill(dt);

            listaNotas.DataSource = dt;
            da.Dispose();

            CargaGrillaNotas();
            lblmensaje.Text = "SE ENCONTRARON " + listaNotas.RowCount + " REGISTROS";
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.Close();


        }

        private void listaNotas_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void listaNotas_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                int fila = 0;
                string numero = "";
                string fecha = "";
                string[] xfecha = null;
                metroTabControl1.SelectedTab = metroTabPage2;
                fila = listaNotas.CurrentCell.RowIndex;
                numero = listaNotas.Rows[fila].Cells[0].Value.ToString();
                fecha = listaNotas.Rows[fila].Cells[1].Value.ToString();
                xfecha = fecha.Split(' ');

                Ventas venta = new Ventas();
                DataTable dt = venta.GetDetalleByNumeroCaja("00", numero, xfecha[0]);

                gdrollo.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                gdrollo.RowTemplate.Height = 130;
                gdrollo.AllowUserToAddRows = false;


                gdrollo.DataSource = dt;
                //DataGridViewImageColumn image = new DataGridViewImageColumn();
                //image = (DataGridViewImageColumn)gdrollo.Columns[0];
                //image.ImageLayout = DataGridViewImageCellLayout.Zoom;


                CargaGrillaRollo();
                btnInsertar.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:" + ex.Message);
            }
        }

        private void CargaGrillaRollo()
        {
            //metroTabControl1.SelectedTab = metroTabPage3;
            double neto = 0;
            double iva = 0;
            double total = 0;
            double subtotal = 0;
            int cuenta = 0;

            gdrollo.Columns[0].Width = 80;
            gdrollo.Columns[1].Width = 260;
            gdrollo.Columns[2].Width = 80;
            gdrollo.Columns[3].Width = 80;
            gdrollo.Columns[4].Width = 80;

            gdrollo.Columns[3].DefaultCellStyle.Format = "C0";
            gdrollo.Columns[4].DefaultCellStyle.Format = "C0";


            foreach (DataGridViewRow row in gdrollo.Rows)
            {
                row.Height = 80;

                row.Cells[0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                row.Cells[1].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                row.Cells[2].Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                row.Cells[3].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                row.Cells[4].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;


                row.Cells[1].Style.Font = new Font(row.Cells[1].InheritedStyle.Font, FontStyle.Bold);
                row.Cells[2].Style.Font = new Font(row.Cells[2].InheritedStyle.Font, FontStyle.Bold);
                row.Cells[3].Style.Font = new Font(row.Cells[3].InheritedStyle.Font, FontStyle.Bold);
                row.Cells[4].Style.Font = new Font(row.Cells[4].InheritedStyle.Font, FontStyle.Bold);

                // row.Cells[1].Style.Font = new Font(row.Cells[1].InheritedStyle.Font, FontStyle.Bold);
                row.Cells[2].Style.Font = new Font("Verdana", 10F, FontStyle.Bold);
                row.Cells[3].Style.Font = new Font("Verdana", 10F, FontStyle.Bold);
                row.Cells[4].Style.Font = new Font("Verdana", 10F, FontStyle.Bold);


                subtotal = subtotal + (double)row.Cells[4].Value;
                cuenta = cuenta + 1;


            }
            lblcuentaitems.Text = "ITEMS: " + cuenta.ToString();

            neto = Math.Round(subtotal / 1.19);
            iva = subtotal - neto;
            total = neto + iva;

            //iva = Math.Round(neto * 1.19) - neto;
            lblneto.Text = neto.ToString("C1");
            total = iva + neto;
            lbliva.Text = iva.ToString("C1");
            lbltotal.Text = total.ToString("C1");

            //finalneto.Text = lblneto.Text;
            //finaliva.Text = lbliva.Text;
            //finaltotal.Text = lbltotal.Text;


            gdrollo.ColumnHeadersVisible = false;
            gdrollo.Refresh();
            if (cuenta == 32)
            {
                MetroMessageBox.Show(this, "HA LLEGADO AL LIMITE DE ITEMS POR NOTA. FINALICE PARA CONTINUAR.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            }


        }

        private void btnInsertar_Click(object sender, EventArgs e)
        {

            //int cantidad = Convert.ToDouble(txtcantidad.Text.Replace("$", ""));

            // if (cantidad > 0)
            // {
            //     precio = Convert.ToDouble(lblPrecio.Text.Replace("$", ""));
            //     total = cantidad * precio;

            //     Rollo rollo = new Rollo(1, cantidad, lblcodigo.Text, lbldescripcion.Text, 0, precio, total, "", "", "");
            //     rollo.GrabaItemEnRollo();
            //     rollo.GrabaObservacion(lblcodigo.Text, Funciones.G_CAJA, txtobservacion.Text);
            // }

            double precio = 0;
            double cantidad = 0;
            double total = 0;
            string codigo = "";
            string descripcion = "";
            foreach (DataGridViewRow row in gdrollo.Rows)
            {
                codigo = row.Cells[0].Value.ToString();
                descripcion = row.Cells[1].Value.ToString();
                cantidad = Convert.ToDouble(row.Cells[2].Value);
                precio = Convert.ToDouble(row.Cells[3].Value);
                total = cantidad * precio;
                Rollo ro = new Rollo(1, cantidad, codigo, descripcion, 0, precio, total, "", "", "");
                ro.GrabaItemEnRollo();

            }
            MetroMessageBox.Show(this, "PRODUCTOS AGREGADOS SATISFACTORIAMENTE.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            this.Close();
        }
    }
   
}
