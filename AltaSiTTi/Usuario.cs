using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AltaSiTTi
{
    public partial class Usuario : Form
    { 
        public static String servidor = "148.223.153.43\\MSSQLSERVER1";
        public static String nombre_bd = "bd_SiTTi";
        public static String userbd = "sa";
        public static String passbd = "At3n4";
        public static string nsql = "server=" + servidor + "; database=" + nombre_bd + " ;User ID=" + userbd + ";Password=" + passbd + "; integrated security = false ; MultipleActiveResultSets=True";

        //num empleado tabla empleados
        public static int no_empleado;
        //id empleado tabla usuarios
        public static int id_empleado;

        public static string id_base;
        public static string id_cc;
        public static string id_puesto;
        public Usuario()
        {
            InitializeComponent();
        }

        public static String generarPass()
        {
            string letrasmin = "abcdefghijklmnopqrstuvwxyz";
            string letrasmay = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string numeros = "1234567890";
            string caracteres = "&$#+*";

            string password = "";

            int cc = 0, cn = 0, ce = 0, cm = 0;

            Random r = new Random();

            while (password.Length < 7)
            {
                switch (r.Next(0, 4))
                {
                    case 0:
                        if (cc < 3)
                        {
                            char c = letrasmin[r.Next(letrasmin.Length)];
                            cc++;
                            password += c;
                        }
                        break;
                    case 1:
                        if (cm < 2)
                        {
                            char mm = letrasmay[r.Next(letrasmay.Length)];
                            cm++;
                            password += mm;
                        }

                        break;
                    case 2:
                        if (cn < 1)
                        {
                            char n = numeros[r.Next(numeros.Length)];
                            cn++;
                            password += n;
                        }

                        break;
                    case 3:
                        if (ce < 1)
                        {
                            char eee = caracteres[r.Next(caracteres.Length)];
                            ce++;
                            password += eee;
                        }
                        break;
                }
            }
            return password;
        }

        private void Usuario_Load(object sender, EventArgs e)
        {
            //cc
            using (SqlConnection conexion = new SqlConnection(nsql))
            {
                conexion.Open();
                String consulta = "SELECT id_cc,nombre FROM [bd_SiTTi].[dbo].[cg_cc] where status='A' ORDER BY nombre ASC";
                SqlCommand comm = new SqlCommand(consulta, conexion);
                SqlDataReader nwReader = comm.ExecuteReader();
                cc.DisplayMember = "Text";
                cc.ValueMember = "Value";
                while (nwReader.Read())
                {
                    cc.Items.Add(new { Text = nwReader["nombre"].ToString(), Value = nwReader["id_cc"].ToString() });
                }
            }
            
            //base
            using (SqlConnection conexion = new SqlConnection(nsql))
            {
                conexion.Open();
                String consulta = "SELECT id_base,nombre FROM [bd_SiTTi].[dbo].[cg_base] where status='A' Order by nombre ASC";
                SqlCommand comm = new SqlCommand(consulta, conexion);
                SqlDataReader nwReader = comm.ExecuteReader();
                base_nombre.DisplayMember = "Text";
                base_nombre.ValueMember = "Value";
                while (nwReader.Read())
                {
                    base_nombre.Items.Add(new { Text = nwReader["nombre"].ToString(), Value = nwReader["id_base"].ToString() });
                }
            }

            //puesto
            using (SqlConnection conexion = new SqlConnection(nsql))
            {
                conexion.Open();
                String consulta = "SELECT id_puesto,nombre FROM [bd_SiTTi].[dbo].[cg_puesto] WHERE status='A' ORDER BY nombre ASC";
                SqlCommand comm = new SqlCommand(consulta, conexion);
                SqlDataReader nwReader = comm.ExecuteReader();
                puesto.DisplayMember = "Text";
                puesto.ValueMember = "Value";
                while (nwReader.Read())
                {
                    puesto.Items.Add(new { Text = nwReader["nombre"].ToString(), Value = nwReader["id_puesto"].ToString() });
                }
            }
            
            //empresa
            using (SqlConnection conexion = new SqlConnection(nsql))
            {
                conexion.Open();
                String consulta = "SELECT id_empresa,nombre FROM [bd_SiTTi].[dbo].[cg_empresa] where status='A' ORDER BY nombre ASC";
                SqlCommand comm = new SqlCommand(consulta, conexion);
                SqlDataReader nwReader = comm.ExecuteReader();
                empresa.DisplayMember = "Text";
                empresa.ValueMember = "Value";
                while (nwReader.Read())
                {
                    empresa.Items.Add(new { Text = nwReader["nombre"].ToString(), Value = nwReader["id_empresa"].ToString() });
                }
            }


            foreach (String[] empleado in Form1.lista)
            {
                String nmayus = empleado[0].ToString().ToUpper();
                if (nmayus.Contains(Form1.numerodeempelado))
                {
                    nombre.Text = empleado[1];
                    appaterno.Text = empleado[2];
                    apmaterno.Text = empleado[3];
                    cc.Text = empleado[5];
                    base_nombre.Text = empleado[6];
                    numempleado.Text = empleado[0];
                    puesto.Text = empleado[7];
                    correo.Text = empleado[4];
                    ext.Text = empleado[9];
                    telefeno.Text = empleado[8];
                    empresa.Text = empleado[10];
                    user.Text = empleado[1].Substring(0, 1) + empleado[2] + empleado[3].Substring(0, 1);
                    pass.Text = generarPass();
                }
            }
            //Console.WriteLine(Form1.numerodeempelado);
            actividad.SelectedIndex = 3;
            no_empleado = Convert.ToInt32(numempleado.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(String.IsNullOrEmpty(correo.Text))
            {
                MessageBox.Show("El campo correo no puede quedar vacio.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //buscar en tabla empleados
            using (SqlConnection conexion = new SqlConnection(nsql))
            {
                conexion.Open();
                String consulta = "SELECT * FROM [bd_SiTTi].[dbo].[cg_empleado] where status = 'A' AND no_empleado='"+ no_empleado + "'";
                SqlCommand comm = new SqlCommand(consulta, conexion);
                SqlDataReader nwReader = comm.ExecuteReader();
                while (nwReader.Read())
                {
                    id_empleado = Convert.ToInt32(nwReader["id_empleado"].ToString());
                    AltaSiTTi();
                    return;
                }
            }
            
            
            //insertar usuario en tabla empelados
            using (SqlConnection conexion2 = new SqlConnection(nsql))
            {
                try
                {
                    String insert = "INSERT INTO cg_empleado VALUES ('" + no_empleado + "' ,'" + nombre.Text + "' ,'" + appaterno.Text + "' ,'" + apmaterno.Text + "' ,'" + correo.Text + "' ,'" + telefeno.Text + "' ,'" + ext.Text + "' ,'" + (base_nombre.SelectedItem as dynamic).Value + "' ,'" + (cc.SelectedItem as dynamic).Value + "' ,'" + (puesto.SelectedItem as dynamic).Value + "' ,'N' ,'A' ,'N' ,"+actividad.SelectedIndex+")";
                    Console.WriteLine(insert);
                    conexion2.Open();
                    SqlCommand comm2 = new SqlCommand(insert, conexion2);
                    comm2.ExecuteReader();


                    using (SqlConnection conexion = new SqlConnection(nsql))
                    {
                        conexion.Open();
                        String consulta = "SELECT id_empleado FROM [bd_SiTTi].[dbo].[cg_empleado] where no_empleado='" + no_empleado + "'";
                        SqlCommand comm = new SqlCommand(consulta, conexion);
                        SqlDataReader nwReader = comm.ExecuteReader();
                        while (nwReader.Read())
                        {
                            id_empleado = Convert.ToInt32(nwReader["id_empleado"].ToString());
                        }
                    }
                    AltaSiTTi();
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Error al insertar usuario en tabla empleados..\n\nMensaje: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void AltaSiTTi()
        {
            //buscar en tabla usuarios
            using (SqlConnection conexion = new SqlConnection(nsql))
            {
                conexion.Open();
                String consulta = "SELECT * FROM [bd_SiTTi].[dbo].[cg_usuario] where status='A' AND id_empleado='" + id_empleado + "'";
                Console.WriteLine(consulta);
                SqlCommand comm = new SqlCommand(consulta, conexion);
                SqlDataReader nwReader = comm.ExecuteReader();
                while (nwReader.Read())
                {
                    DialogResult dr = MessageBox.Show("El usuario ya cuenta con usuario sitti\n\nQuieres reenviar correo con usuario y contraseña?", "Correo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes)
                    {
                        user.Text = nwReader["nick"].ToString();
                        pass.Text = nwReader["pass"].ToString();
                        enviarcorreo(); 
                        MessageBox.Show("Datos enviados.", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    Close();
                    return;
                }
                
                //insertar usuario en tabla empelados
                using (SqlConnection conexion2 = new SqlConnection(nsql))
                {
                    try
                    {
                        String insert = "INSERT INTO cg_usuario VALUES ('" + id_empleado + "' ,'" + user.Text + "' ,'" + pass.Text + "' ,'U' ,'1' ,'A')";
                        conexion2.Open();
                        SqlCommand comm2 = new SqlCommand(insert, conexion2);
                        comm2.ExecuteReader();

                        DialogResult dr = MessageBox.Show("Enviar correo con usuario y contraseña?", "Correo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dr == DialogResult.Yes)
                        {
                            enviarcorreo();
                        }
                        MessageBox.Show("Usuario creado con exito.", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al insertar usuario en tabla usuarios.\n\nMensaje: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void label13_Click(object sender, EventArgs e)
        {
            pass.Text = generarPass();
        }

        private void enviarcorreo()
        {
            List<string> lstAllRecipients = new List<string>();
            //Below is hardcoded - can be replaced with db data
            lstAllRecipients.Add(correo.Text);

            Outlook.Application outlookApp = new Outlook.Application();
            Outlook._MailItem oMailItem = (Outlook._MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            Outlook.Inspector oInspector = oMailItem.GetInspector;

            // Recipient
            Outlook.Recipients oRecips = (Outlook.Recipients)oMailItem.Recipients;
            foreach (String recipient in lstAllRecipients)
            {
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(recipient);
                oRecip.Resolve();
            }

            oMailItem.Subject = "Usuario SiTTi";


            oMailItem.Attachments.Add(System.IO.Directory.GetCurrentDirectory()+"\\Manual Sitti V2.0.pdf", Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);


            String FirmaBody = oMailItem.HTMLBody;

            //Body para Edson
            oMailItem.HTMLBody = "<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta http-equiv=Content-Type content=\"text/html; charset=iso-8859-1\"><meta name=Generator content=\"Microsoft Word 15 (filtered medium)\"><!--[if !mso]><style>v:* {behavior:url(#default#VML);} o:* {behavior:url(#default#VML);} w:* {behavior:url(#default#VML);} .shape {behavior:url(#default#VML);} </style><![endif]--><style><!-- /* Font Definitions */ @font-face 	{font-family:Helvetica; 	panose-1:2 11 6 4 2 2 2 2 2 4;} @font-face 	{font-family:\"Cambria Math\"; 	panose-1:2 4 5 3 5 4 6 3 2 4;} @font-face 	{font-family:Calibri; 	panose-1:2 15 5 2 2 2 4 3 2 4;} @font-face 	{font-family:Tahoma; 	panose-1:2 11 6 4 3 5 4 4 2 4;} /* Style Definitions */ p.MsoNormal, li.MsoNormal, div.MsoNormal 	{margin:0cm; 	margin-bottom:.0001pt; 	font-size:11.0pt; 	font-family:\"Calibri\",sans-serif; 	mso-fareast-language:EN-US;} a:link, span.MsoHyperlink 	{mso-style-priority:99; 	color:#0563C1; 	text-decoration:underline;} a:visited, span.MsoHyperlinkFollowed 	{mso-style-priority:99; 	color:#954F72; 	text-decoration:underline;} p.msonormal0, li.msonormal0, div.msonormal0 	{mso-style-name:msonormal; 	mso-margin-top-alt:auto; 	margin-right:0cm; 	mso-margin-bottom-alt:auto; 	margin-left:0cm; 	font-size:12.0pt; 	font-family:\"Times New Roman\",serif;} span.EstiloCorreo18 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:windowtext;} span.EstiloCorreo19 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo20 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo21 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo22 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo23 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo24 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo25 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo26 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo27 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo28 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo29 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo30 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo31 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo32 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo33 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo34 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo35 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo36 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo37 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F3864;} span.EstiloCorreo38 	{mso-style-type:personal; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} span.EstiloCorreo39 	{mso-style-type:personal-reply; 	font-family:\"Calibri\",sans-serif; 	color:#1F497D;} .MsoChpDefault 	{mso-style-type:export-only; 	font-size:10.0pt;} @page WordSection1 	{size:612.0pt 792.0pt; 	margin:70.85pt 3.0cm 70.85pt 3.0cm;} div.WordSection1 	{page:WordSection1;} --></style><!--[if gte mso 9]><xml> <o:shapedefaults v:ext=\"edit\" spidmax=\"1026\" /> </xml><![endif]--><!--[if gte mso 9]><xml> <o:shapelayout v:ext=\"edit\"> <o:idmap v:ext=\"edit\" data=\"1\" /> </o:shapelayout></xml><![endif]--></head><body lang=ES-MX link=\"#0563C1\" vlink=\"#954F72\"><div class=WordSection1><p class=MsoNormal><span style='color:#2F5597;mso-fareast-language:ES-MX'>Buen día.</span><span style='color:#2F5597'><o:p></o:p></span></p><p class=MsoNormal><span style='color:#2F5597;mso-fareast-language:ES-MX'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span style='color:#2F5597;mso-fareast-language:ES-MX'>Por medio de la presente te hacemos llegar tu usuario y contraseña para el uso del Sistema de Tickets SiTTi. Es muy importante que utilices correctamente esta información considerando los siguientes aspectos:<o:p></o:p></span></p><p class=MsoNormal><span style='color:#2F5597;mso-fareast-language:ES-MX'>&#8226;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; El acceso al SiTTi es a través de la dirección <a href=\"http://148.223.153.43/SiTTi\"><span style='color:#2F5597'>http://148.223.153.43/SiTTi</span></a><o:p></o:p></span></p><p class=MsoNormal><span style='color:#2F5597;mso-fareast-language:ES-MX'>&#8226;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; La utilización del usuario y contraseña son de uso PERSONAL, es decir, que tú eres el responsable de cualquier solicitud que se haga en el sistema con ese usuario, por lo tanto no debes prestarlo a nadie.<o:p></o:p></span></p><p class=MsoNormal><span style='mso-fareast-language:ES-MX'><o:p>&nbsp;</o:p></span></p><table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse'><tr><td width=123 style='width:91.9pt;border:solid windowtext 1.0pt;background:#A8D08D;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='mso-fareast-language:ES-MX'>USUARIO SITTI<o:p></o:p></span></b></p></td><td width=123 style='width:92.15pt;border:solid windowtext 1.0pt;border-left:none;background:#A8D08D;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b><span style='mso-fareast-language:ES-MX'>CONTRASEÑA<o:p></o:p></span></b></p></td></tr><tr><td width=123 style='width:91.9pt;border:solid windowtext 1.0pt;border-top:none;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><span style='color:#1F497D;mso-fareast-language:ES-MX'>" + user.Text + "</span><span style='mso-fareast-language:ES-MX'><o:p></o:p></span></p></td><td width=123 style='width:92.15pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><span style='color:#1F497D'>" + pass.Text + "<o:p></o:p></span></p></td></tr></table><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><b><span style='color:#2F5597;mso-fareast-language:ES-MX'>Anexo liga al Directorio Telefónico de Corporativo UNNE:</span></b><b><span style='mso-fareast-language:ES-MX'> </span></b><span style='mso-fareast-language:ES-MX'><a href=\"http://148.223.153.37:8080/\">http://148.223.153.37:8080/</a><o:p></o:p></span></p>";

            oMailItem.HTMLBody += FirmaBody;
            oMailItem.Display(true);
        }

        #region Copiar al Portapeles
        private void label1_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(nombre.Text, true);
        }

        private void label2_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(appaterno.Text, true);
        }

        private void label3_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(apmaterno.Text, true);
        }

        private void label4_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(cc.Text, true);
        }

        private void label5_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(base_nombre.Text, true);
        }

        private void label6_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(numempleado.Text, true);
        }

        private void label11_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(empresa.Text, true);
        }

        private void label7_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(puesto.Text, true);
        }

        private void label8_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(correo.Text, true);
        }

        private void pass_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(pass.Text, true);
        }

        private void user_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(user.Text, true);
        }
        #endregion
    }
}
