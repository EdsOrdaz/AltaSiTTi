using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AltaSiTTi
{
    public partial class Form1 : Form
    {
        //0.1
        private String versiontext = "0.26";
        private String version = "261943f3a93b683ceeac658927f3923f";
        public static String conexionsqllast = "server=148.223.153.37,5314; database=InfEq;User ID=eordazs;Password=Corpame*2013; integrated security = false ; MultipleActiveResultSets=True";

        public static String servivor = "40.76.105.1,5055";
        public static String basededatos = "Nom2001";
        public static String usuariobd = "reportesUNNE";
        public static String passbd = "8rt=h!RdP9gVy";
        public static string nsql = "server=" + servivor + "; database=" + basededatos + " ;User ID=" + usuariobd + ";Password=" + passbd + "; integrated security = false ; MultipleActiveResultSets=True";


        public static List<String[]> lista = new List<String[]>();
        public static String nombreusuario;
        public static String numerodeempelado;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection conexion2 = new SqlConnection(conexionsqllast))
                {
                    conexion2.Open();
                    String sql2 = "SELECT (SELECT valor FROM Configuracion WHERE nombre='ASitti_Version') as version,valor FROM Configuracion WHERE nombre='ASitti_hash'";
                    SqlCommand comm2 = new SqlCommand(sql2, conexion2);
                    SqlDataReader nwReader2 = comm2.ExecuteReader();
                    if (nwReader2.Read())
                    {
                        if (nwReader2["valor"].ToString() != version)
                        {
                            MessageBox.Show("No se puede inciar porque hay una nueva version.\n\nNueva Version: " + nwReader2["version"].ToString() + "\nVersion actual: " + versiontext + "\n\nEl programa se cerrara.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Application.Exit();
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se puedo verificar la version del programa.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en validar la version del programa\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }
            buscar.Enabled = false;
            if (backgroundWorker1.IsBusy != true)
            {
                lista.Clear();
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                using (SqlConnection conexion = new SqlConnection(nsql))
                {
                    conexion.Open();
                    String consulta = "SELECT nomtrab.nombre, nomtrab.apepat, nomtrab.apemat, ISNULL(CC.desubi,'') AS [cc], ISNULL(nompais.despai,'') AS [base], nomtrab.cvetra as numempleado, ISNULL(nomciud.desciu,'') AS [empresa], RTRIM(LTRIM(ISNULL(nompues.despue,''))) AS [Puesto], nomtrab.email, nomtrab.tel1, nomtrab.numext FROM [Nom2001].[dbo].[nomtrab]  LEFT JOIN nomubic CC ON nomtrab.cvepa2 = CC.cvepai AND nomtrab.cveci2 = CC.cveciu AND nomtrab.cvecia = CC.cvecia AND nomtrab.cveubi = CC.cveubi LEFT JOIN nompais ON nomtrab.cvepa2 = nompais.cvepai AND nomtrab.cvecia = nompais.cvecia LEFT JOIN nomciud ON nomtrab.cvepa2 = nomciud.cvepai AND nomtrab.cveci2 = nomciud.cveciu AND nomtrab.cvecia = nomciud.cvecia LEFT JOIN nompues ON nomtrab.cvepue = nompues.cvepue AND nomtrab.cvecia = nompues.cvecia where status='A'";
                    SqlCommand comm = new SqlCommand(consulta, conexion);
                    SqlDataReader nwReader = comm.ExecuteReader();
                    while (nwReader.Read())
                    {
                        String[] n = new String[12];
                        n[0] = nwReader["numempleado"].ToString().TrimEnd(' ');
                        n[1] = nwReader["nombre"].ToString().TrimEnd(' ');
                        n[2] = nwReader["apepat"].ToString().TrimEnd(' ');
                        n[3] = nwReader["apemat"].ToString().TrimEnd(' ');
                        n[4] = nwReader["email"].ToString().TrimEnd(' ');
                        n[5] = nwReader["cc"].ToString().TrimEnd(' ');
                        n[6] = nwReader["base"].ToString().TrimEnd(' ');
                        n[7] = nwReader["Puesto"].ToString().TrimEnd(' ');
                        n[8] = nwReader["tel1"].ToString().TrimEnd(' ');
                        n[9] = nwReader["numext"].ToString().TrimEnd(' ');
                        n[10] = nwReader["empresa"].ToString().TrimEnd(' ');
                        n[11] = n[1] + " " + n[2] + " " + n[3];
                        lista.Add(n);
                    }
                }
            }
            catch (System.Exception ex)
            {
                e.Cancel = true;
                MessageBox.Show("Error al cargar lista de empelados\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if(lista.Any())
            {
                foreach (String[] n in lista)
                {
                    dataGridView1.Rows.Add(n[11], n[0]);
                }
            }
            buscar.Enabled = true;
            buscar.Focus();
        }

        private void buscar_KeyUp(object sender, KeyEventArgs e)
        {
            dataGridView1.Rows.Clear();
            if (!string.IsNullOrEmpty(buscar.Text))
            {
                foreach (String[] empleado in lista)
                {
                    String nmayus = empleado[11].ToString().ToUpper();
                    String nemple = empleado[0].ToString().ToUpper();
                    if (nmayus.Contains(buscar.Text.ToUpper()) || nemple.Contains(buscar.Text.ToUpper()))
                    {
                        dataGridView1.Rows.Add(empleado[11], empleado[0]);
                    }
                }
            }
            else
            {
                foreach (String[] empleado in lista)
                {
                    dataGridView1.Rows.Add(empleado[11], empleado[0]);
                }
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            nombreusuario = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(dataGridView1.Rows[e.RowIndex].Cells["nombre"].Value.ToString().ToLower());
            Clipboard.SetDataObject(nombreusuario, true);
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            nombreusuario = dataGridView1.Rows[e.RowIndex].Cells["nombre"].Value.ToString();
            numerodeempelado = dataGridView1.Rows[e.RowIndex].Cells["num"].Value.ToString();
            Clipboard.SetDataObject(nombreusuario, true);
            Usuario usuario = new Usuario();
            usuario.ShowDialog();
        }
    }
}
