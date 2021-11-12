using MetroFramework.Controls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using NationalInstruments.Visa;
using System.IO;
using Ivi.Visa;
using MetroFramework;
using System.Threading;

namespace AutocalMyp
{
    public partial class AutoCal : MetroFramework.Forms.MetroForm
    {

        public List<_instrumenta> list_insta = new List<_instrumenta>(); // Generador
        public List<_instrumentb> list_instb = new List<_instrumentb>(); // Lectura         
        public int i = 0;

        public static _excel excel = new _excel();

        public AutoCal()
        {
            InitializeComponent();
        }

        private void AutoCal_Load(object sender, EventArgs e)
        {
            metroTextBox_informe.Enabled = false;
            metroButton_iniciar.Enabled = false;
            resorce();
        }

        public void resorce()
        {
            try
            {
                IEnumerable<string> devices;
                var rm = new ResourceManager();
                var resource = rm.Find("?*");
                devices = resource;
                foreach (string value in devices)
                {
                    metroCombodevices.Items.Add(value);
                }
            }
            catch (Exception er)
            {
                MetroMessageBox.Show(this, "Alerta!.", "Mensaje de notificación \r\n" + er.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }

        private void metroCombodevices_SelectedValueChanged(object sender, EventArgs e)
        {
            string device = metroCombodevices.SelectedItem.ToString();
            string result = conneccionvisa(device);
            metroLabel_idn.Text = result.ToString();
        }

        public string conneccionvisa(string device)
        {
            string res;

            try
            {
                var mydevice = openVisa(device);
                writeVisa(mydevice, "*IDN?\n");
                res = readVisa(mydevice);
                closeVisa(mydevice);
                return res;
            }
            catch (Exception ex)
            {
                return res = "Error:101" + ex; //Error 101 Error valor:null
            }

        }

        public void showlista()
        {
            var n = list_insta.Count - 1;
            for (int y = n; y < list_insta.Count; y++)
            {
                i++;
                MetroTile instr = new MetroTile();
                instr.Name = "lector_" + i;
                instr.Width = 316;
                instr.Height = 75;
                instr.Dock = DockStyle.Top;
                instr.Text = "Instrumento:" + list_insta[y].modelo.ToString() + "\n Resource:" + list_insta[y].resource.ToString() + "\n #Informe:" + list_insta[y].informe.ToString() + "\n Modo:" + list_insta[y].modo.ToString();
                instr.TileCount = i;
                instr.TextAlign = ContentAlignment.TopLeft;
                instr.Style = MetroFramework.MetroColorStyle.Green;
                metroPanel1.Controls.Add(instr);
            }
        }

        public void showlistb()
        {
            var n = list_instb.Count - 1;
            for (int y = n; y < list_instb.Count; y++)
            {
                i++;
                MetroTile instr = new MetroTile();
                instr.Name = "generador_" + i;
                instr.Width = 316;
                instr.Height = 75;
                instr.Dock = DockStyle.Top;
                instr.Text = "Instrumento:" + list_instb[y].modelo.ToString() + "\n Resource:" + list_instb[y].resource.ToString() + "\n #Informe:" + list_instb[y].informe.ToString() + "\n Modo:" + list_instb[y].modo.ToString();
                instr.TileCount = i;
                instr.TextAlign = ContentAlignment.TopLeft;
                instr.Style = MetroFramework.MetroColorStyle.Blue;
                metroPanel1.Controls.Add(instr);
            }
        }

        private void metroRadioButton_p_CheckedChanged(object sender, EventArgs e)
        {
            metroTextBox_informe.Text = "";
            metroTextBox_informe.Enabled = false;
        }

        private void metroRadioButton_bc_CheckedChanged(object sender, EventArgs e)
        {
            metroTextBox_informe.Text = "";
            metroTextBox_informe.Enabled = true;
        }

        private void metroComboBox_tipocal_SelectedIndexChanged(object sender, EventArgs e)
        {
            metroButton_iniciar.Enabled = true;
        }

        private void metroButton_guardar_Click(object sender, EventArgs e)
        {
            string[] arraymodo = { "Generador", "Lector" };
            string resource = metroCombodevices.SelectedItem.ToString();
            string nombre = metroLabel_idn.Text.ToString();
            var dataidn = nombre.Split(',');
            nombre = dataidn[0] + "," + dataidn[1];

            string informe = metroTextBox_informe.Text.ToString();

            bool equipo_bc = metroRadioButton_bc.Checked;
            bool equipo_p = metroRadioButton_patron.Checked;

            bool equipo_gen = metroRadioButton_generador.Checked;
            bool equipo_lec = metroRadioButton_lector.Checked;

            if (equipo_gen == true)
            {
                if (equipo_p == true)
                {
                    informe = "";
                    list_insta.Add(new _instrumenta { modelo = nombre, resource = resource, informe = informe, modo = arraymodo[0], device = openVisa(resource), file = "" });
                    showlista();
                }
                else
                {
                    if (informe != "")
                    {
                        var data = nombre.Split(',');
                        list_insta.Add(new _instrumenta { modelo = nombre, resource = resource, informe = informe, modo = arraymodo[0], device = openVisa(resource), file = informe + "_" + data[1] });
                        showlista();
                    }
                    else
                    {
                        MetroMessageBox.Show(this, "Alerta! Campo de informe vacio.", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else if (equipo_lec == true)
            {
                if (equipo_p == true)
                {
                    informe = "";
                    list_instb.Add(new _instrumentb { modelo = nombre, resource = resource, informe = informe, modo = arraymodo[1], device = openVisa(resource), file = "" });
                    showlistb();
                }
                else
                {
                    if (informe != "")
                    {
                        var data = nombre.Split(',');
                        list_instb.Add(new _instrumentb { modelo = nombre, resource = resource, informe = informe, modo = arraymodo[1], device = openVisa(resource), file = informe + "_" + data[1] });                        
                        showlistb();
                    }
                    else
                    {
                        MetroMessageBox.Show(this, "Alerta! Campo de informe vacio.", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            this.Clear();
        }

        private void metroButton_iniciar_Click(object sender, EventArgs e)
        {
            //Validar select de tipo de calibración
            string tipo_cal = metroComboBox_tipocal.SelectedIndex.ToString();
            if (list_insta.Count > 0 && list_instb.Count > 0)
            {
                if (int.Parse(tipo_cal) == 0)
                {//Calibración de Fuentes
                    foreach (_instrumenta data in list_insta)
                    {
                        Createfile(data.file);
                    }
                }
                else if (int.Parse(tipo_cal) == 1)
                { // Calibración de Multimetros
                    foreach (_instrumentb data in list_instb)
                    {
                        var modelo = data.modelo.Split(',');
                        Createfile(data.informe + "_" + modelo[1].ToString());
                    }
                }
                //Validacion de la creacion de archivos- pendiente
                inicio_secuencia(tipo_cal);
            }
            else
            {
                string alerta = "Alerta!";
                if (list_insta.Count == 0)
                {
                    alerta += " No se ha encontrado equipo Generador.";
                }
                if (list_instb.Count == 0)
                {
                    alerta += " No se ha encontrado equipo lector.";
                }

                MetroMessageBox.Show(this, alerta, "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public string inicio_secuencia(string funcion)
        {
            if (int.Parse(funcion) == 0)
            { // Calibracion de fuente
                //CalibracionFuente
                try
                {
                    CalibracionFuente();
                    //metroTextBox_command.Text += "\r\n" + "Fin de la calibracion:" + DateTime.Now.ToString();
                    MetroMessageBox.Show(this, "Calibración exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                   // metroTextBox_command.Text += "\r\n" + "Fin de la calibracion:" + DateTime.Now.ToString();
                    MetroMessageBox.Show(this, "Error!" + ex.ToString(), "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }
            else if (int.Parse(funcion) == 1)
            { //Calibracion de multimetros
                try
                {
                    CalibracionMultimetro();
                    //listBox_command.Items.Add(":" + "Fin de la calibracion:" + DateTime.Now.ToString());
                   // metroTextBox_command.Text += "\r\n" + "Fin de la calibracion:" + DateTime.Now.ToString();
                    MetroMessageBox.Show(this, "Calibración exitosamente!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                   // metroTextBox_command.Text += "\r\n" + "Fin de la calibracion:" + DateTime.Now.ToString();
                    MetroMessageBox.Show(this, "Error!" + ex.ToString(), "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            int n_generadores = list_insta.Count;
            for (int a = 0; a < n_generadores; a++)
            {
                closeVisa(list_insta[a].device);
            }

            int n_lectores = list_instb.Count;
            for (int a = 0; a < n_lectores; a++)
            {
                closeVisa(list_instb[a].device);
            }

            return "";
        }

        #region  Funciones de calibracion  

        public void CalibracionFuente()
        {

           // metroTextBox_command.Text += "\r\n" + "#001 Inicio de calibracion:" + DateTime.Now.ToString();
            #region Valores de las funciones
            string[] nombrefunc = new string[] { "DC V", "DC I" };
            int[] rangoxFunc = new int[] { 2, 2 };
            string[] vmedir = new string[] { "2,2", "2,2" };
            string[] tiempovm = new string[] { "2000,2000", "2000,2000" };

            int[] funcactiva = new int[] { 1, 1 };
            int[] modocal = new int[] { 2, 2 }; // Parelelo 0 | Serie 1 | Separado 2

            /* Equipo de lectura */
            /*  DC V */
            /*Esta configuracion es para la medicion por el metodo indirecto*/
            //Comandos para el equipo lector
            //Agilent 34410A
            string[] commandEqLectura = new string[] { "CONF:VOLT:DC 100 V,CONF:VOLT:DC 100 V", "CONF:VOLT:DC 1 V,CONF:VOLT:DC 1 V" };
            //Fluke 8846a
            //string[] commandEqLectura = new string[] { "CONF:VOLT:DC 100 V,CONF:VOLT:DC 100 V", "CONF:CURR:DC 10,CONF:CURR:DC 10" };
            /* Equipo Generador / Calibrador */
            //string[] commandEqGenerador = new string[] { "VOLTage:RANGe LOW\n Volt 1.5\n Output on\n,VOLTage:RANGe LOW\n Volt 13.5\n Output on\n,VOLTage:RANGe LOW\n Volt 3\n Output on\n,VOLTage:RANGe HIGH \n Volt 27\n Output on\n", "VOLTage:RANGe LOW\n APPL 10,0.7\n Output on\n,VOLTage:RANGe LOW\n APPL 10,6.3\n Output on\n","VOLTage:RANGe HIGH\n APPL 10,0.4\n Output on\n,VOLTage:RANGe HIGH\n APPL 10,3.6\n Output on\n" };

            //Comandos para el equipo generador escritura
            string[] array1 = { "VOLTage:RANGe LOW\n Volt 1.5\n Output on\n", "VOLTage:RANGe LOW\n Volt 13.5\n Output on\n", "VOLTage:RANGe LOW\n Volt 3\n Output on\n", "VOLTage:RANGe HIGH \n Volt 27\n Output on\n" };
            string[] array2 = { "VOLTage:RANGe LOW\n APPL 10,0.7\n Output on\n", "VOLTage:RANGe LOW\n APPL 10,6.3\n Output on\n", "VOLTage:RANGe HIGH\n APPL 10,0.4\n Output on\n", "VOLTage:RANGe HIGH\n APPL 10,3.6\n Output on\n" };


            #endregion
            for (int i = 0; i < funcactiva.Length; i++)
            {
                int nrango = rangoxFunc[i];
                int enable = funcactiva[i];

                string[] vmedir_temp = new string[] { "" };
                vmedir_temp = vmedir[i].Split(',');

                string[] tiempovm_temp = new string[] { "" };
                tiempovm_temp = tiempovm[i].Split(',');

                //string[] configEqGenerador = new string[] { "" };
                //configEqGenerador = commandEqGenerador[i].Split(',');

                string[] configEqLectura = new string[] { "" };
                configEqLectura = commandEqLectura[i].Split(',');

                //myGenerador.WriteString("*CLS", true); //Cleaning configuracion
                writeVisa(list_instb[0].device, "*CLS"); // Este device esta en modo lector


                string funcion = nombrefunc[i];

                if (funcion == "DC V" && enable == 1)
                {
                   metroTextBox_command.Text += "\r\n" + "Funcion | "+ funcion;
                   FunctionReadFuente(i, array1, configEqLectura, nrango, vmedir_temp, tiempovm_temp, modocal[i]);
                }
                if (funcion == "DC I" && enable == 1)
                {
                    metroTextBox_command.Text += "\r\n" + "Pregunta | " + " Mediremos Corriente, cambiar cables a los puertos correctos, cuando estes listo presiona -Yes-?";

                    DialogResult dr = new DialogResult();
                    do
                    {
                        dr = MetroFramework.MetroMessageBox.Show(this,
                      "\n\n Mediremos Corriente, cambiar cables a los puertos correctos, cuando estes listo presiona -Yes-?", "Alerta!",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Question
                        );
                    } while (dr != DialogResult.Yes);

                    if (dr == DialogResult.Yes)
                    {
                        metroTextBox_command.Text += "\r\n" + "Funcion | " + funcion;
                        FunctionReadFuente(i, array2, configEqLectura, nrango, vmedir_temp, tiempovm_temp, modocal[i], 0.1);
                    }
                    else
                    {
                        //Terminar calibracion                     
                    }
                }

            }
        }

        public void FunctionReadFuente(int funcion, string[] configCalibrador, string[] configEquipo, int nrango, string[] arrmediciones, string[] arrtiempovm, int modocal, double ohm = 0)
        {
            string[] nombrevrango = new string[] { "15 V,30 V", "7 A,4 A" }; //DC V, Corriente DC
            string[] nombrevnominal = new string[] { "1.5,13.5,3,27", "0.7,6.3,0.4,3.6" }; //DC V, Corriente DC
            string[] nombremodocal = new string[] { "Paralelo", "Serie", "Separado" };
            string[] nombrefunc = new string[] { "DC V", "DC I" };

            string[] vrango_temp = new string[] { "" };
            vrango_temp = nombrevrango[funcion].Split(',');

            string[] vnominal_temp = new string[] { "" };
            vnominal_temp = nombrevnominal[funcion].Split(',');

            //int n_eqcalibrar = EquipoAddrs.Length;// Leer lista de fuentes modo generador

            int n_equipobc = list_insta.Count;

            string retorno = "";
            int pivoteconfigCalib = 0;
            string configCal = "";
            double[] lectura = new double[3];
            bool lecturacorrecta = true;

            for (int a = 0; a < n_equipobc; a++)
            {
                pivoteconfigCalib = 0;
                if (modocal == 2)
                {
                    if (n_equipobc > 0) // Equipo por equipo, opcion cuando sean mas de uno
                    {
                        metroTextBox_command.Text += "\r\n" + "Pregunta | " + "\n\n Se realizara la calibración del equipo " + list_insta[a].modelo.ToString() + " de la función: " + nombrefunc[funcion] + ", asi que se requiere la conección en " + nombremodocal[modocal];
                        DialogResult dr = new DialogResult();
                        do
                        {
                            dr = MetroFramework.MetroMessageBox.Show(this,
                          "\n\n Se realizara la calibración del equipo " + list_insta[a].modelo.ToString() + " de la función: " + nombrefunc[funcion] + ", asi que se requiere la conección en " + nombremodocal[modocal] + ". Listo para continuar... ", "Alerta!",
                           MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question
                            );
                        } while (dr != DialogResult.Yes);

                    }
                    /* Escribir la funcion por cada equipo activo, en el archivo correspondiente */
                    Writefuncion(list_insta[a].file.ToString(), nombrefunc[funcion].ToString());

                    writeVisa(list_instb[0].device, "*RST");                    
                    writeVisa(list_instb[0].device, "*CLS");
                }

                for (int i = 0; i < nrango; i++) // Numero de rangos por funcion [2,2]
                {
                    int nmediciones = int.Parse(arrmediciones[i]); //numero de medicionees por rango                  
                    string coleccionvpatron = "";
                    string coleccionlecturas = "";

                    for (int j = 0; j < nmediciones; j++) // numero de mediciones que se van hacer en el rango de la funcion. Ejemplo: DC V->15 V-> #2 -> repeticiones 3 por cada nmediciones
                    {
                        //myGenerador.WriteString("*CLS", true);
                        writeVisa(list_insta[a].device, "*CLS");
                        if (j == 0)
                        {
                            //Configuracion para la funcion del equipo por rango, Equipo en modo lector                            
                            writeVisa(list_instb[0].device, "*CLS");
                            metroTextBox_command.Text += "\r\n" + "Wait | " + "1000";
                            Thread.Sleep(1000);
                            string configtemp = configEquipo[i];
                            writeVisa(list_instb[0].device, configtemp);

                            //Aqui se hace la configuracion del equipo conectado en cada rango                         
                            /* Escribir los rangos por cada equipo simultaneamente, en el archivo correspondiente */
                            Writerango(list_insta[a].file.ToString(), vrango_temp[i].ToString());
                            //metroTextBox_command.Text += "\r\n" + vrango_temp[i].ToString();

                        }
                        //Configuracion para el equipo generador
                        configCal = configCalibrador[pivoteconfigCalib];
                        writeVisa(list_insta[a].device, configCal);
                        writeVisa(list_insta[a].device, "*WAI");                        

                        int delay = int.Parse(arrtiempovm[i]);
                        metroTextBox_command.Text += "\r\n" + "Wait | " + delay.ToString();
                        Thread.Sleep(delay);

                        coleccionvpatron = coleccionvpatron + vnominal_temp[pivoteconfigCalib] + ",";

                        for (int k = 0; k < 3; k++)
                        {
                            //Reading...                            
                            writeVisa(list_instb[0].device, "READ?");
                            retorno = readVisa(list_instb[0].device);
                            lectura[k] = Convert.ToDouble(retorno);
                            metroTextBox_lecturas.Text += "\r\n" + lectura[k].ToString();
                            metroTextBox_command.Text += "\r\n" + "Wait | " + "500";
                            Thread.Sleep(500);
                        }

                        if (lecturacorrecta == true)
                        {
                            for (int c = 0; c < lectura.Length; c++)
                            {
                                coleccionlecturas = coleccionlecturas + lectura[c].ToString() + ",";
                            }
                        }
                        pivoteconfigCalib = pivoteconfigCalib + 1;
                    }
                    //metroTextBox_command.Text += "\r\n" + "Guardando lectura ...";

                    Writeread(list_insta[a].file.ToString(), coleccionvpatron, coleccionlecturas);
                }

                for (int c = 0; c < n_equipobc; c++)
                {
                    writeVisa(list_insta[a].device, "*RST");
                }
                writeVisa(list_instb[0].device, "*RST");
                ////myGenerador.WriteString("*RST", true);
                //writeVisa(list_insta[a].device, "*RST");
                metroTextBox_command.Text += "\r\n" + "Wait | " + "2000";
                Thread.Sleep(2000);
            }
            metroTextBox_command.Text += "\r\n" + "Wait | " + "2000";
            Thread.Sleep(2000);
        }

        #endregion 

        #region  Funciones de calibracion  

        public void CalibracionMultimetro()
        {
            //Console.WriteLine("#002 Inicio de calibracion:" + DateTime.Now.ToString()); //                          
            //metroTextBox_command.Text += "\r\n" + "#002 Inicio de calibracion:" + DateTime.Now.ToString();

            #region Valores de las funciones
            //Resistencia, DC V, AC V @60hz, Frecuencia @60hz, Corriente AC @60hz, Corriente DC
            /* Terminos de las funciones, numero de rangos, nombre de los rangos, valores nominales, unidades del rango,valor a medir por rango, tiempo de espera por rango, comandos por punto (calibrador), comandos de configuracion (Equipo) */
            int[] rangoxFunc = new int[] { 7, 5, 5, 2, 2, 4 };
            string[] vmedir = new string[] { "2,2,2,2,2,2,2", "4,4,2,2,2", "2,2,2,2,2", "2,2", "2,2", "2,2,2,2" };
            string[] tiempovm = new string[] { "2500,2500,2500,2500,2500,2500,2500", "3000,2500,2500,4000,6500", "4000,3500,4000,4500,5000", "3000,3000", "2000,2500", "2000,2000,2500,3000" };
            string[] nombrefunc = new string[] { "Resistencia", "DC V", "AC V @60hz", "Frecuencia", "AC I @60hz", "DC I" };
            int[] funcactiva = new int[] { 1, 1, 1, 1, 1, 1 };
            int[] modocal = new int[] { 2, 0, 0, 0, 1, 1 }; // Parelelo 0 | Serie 1 | Separado 2

            //string[] nombrevrango = new string[] { "100 Ω,1 kΩ,10 kΩ,100 kΩ,1 MΩ,10 MΩ,100 MΩ", "100 mV,1 V,10 V,100 V,1000 V", "100 mV,1 V,10 V,100 V,750 V", "10 Hz - 40 Hz,40 Hz - 300 kHz", "1 A,3 A", "10 mA,100 mA,1 A,3 A" }; //Resistencia, DC V, AC V @60hz, Frecuencia @60hz, Corriente AC @60hz, Corriente DC
            //string[] nombrevnominal = new string[] { "10,90,0.1,0.9,1,9,10,90,100,900,100,900,100,900", "10,90,0.1,0.9,1,9,10,90,100,900", "10,90,0.1,0.9,1,9,10,90,75,675", "10,40,40,300", "0.1,0.9,0.3,2.7", "1,9,10,90,0.1,0.9,0.3,2.7" }; //Resistencia, DC V, AC V @60hz, Frecuencia @60hz, Corriente AC @60hz, Corriente DC
            /* revisar lo de las unidades*/
            //string[] unidad = new string[] { "Ω", "V", "V", "Hz", "A", "A" }; //Resistencia, DC V, AC V @60hz, Frecuencia @60hz, Corriente AC @60hz, Corriente DC
            #endregion
            #region Terminos de las funciones
            /*  Resistencia */
            string[] configCalibradorRest = new string[] {
                "r164/O10/S0<CR>",
                "r164/O90/S0<CR>",
                "r165/O100/S0<CR>",
                "r166/O0.9/S0<CR>",
                "r167/O1/S0<CR>",
                "r168/O9/S0<CR>",
                "r169/O10/S0<CR>",
                "r170/O90/S0<CR>",
                "r171/O100/S0<CR>",
                "r172/O0.9/S0<CR>",
                "r173/O1/S0<CR>",
                "r174/O9/S0<CR>",
                "r175/O10/S0<CR>",
                "r176/O90/S0<CR>" };
            //[2,2,2,2,2,2,2]
            /*  DC V */
            string[] configCalibradorDcV = new string[] { "r1/O10/S0<CR>", "r1/O-10/S0<CR>", "r1/O90/S0<CR>", "r1/O-90/S0<CR>", "r1/O100/S0<CR>", "r1/O-100/S0<CR>", "r2/O0.9/S0<CR>", "r2/O-0.9/S0<CR>", "r2/O1/S0<CR>", "r3/O9/S0<CR>", "r3/O10/S0<CR>", "r4/O90/S0<CR>", "r4/O100/S0<CR>", "r5/O900/S0<CR>" };
            //[4,4,2,2,2]
            /*  AC V  @60 Hz*/
            string[] configCalibradoAcV = new string[] { "r21/O10/F60/S0<CR>", "r21/O90/F60<CR>", "r21/O100/F60/S0<CR>", "r22/O0.9/F60/S0<CR>", "r22/O1/F60/S0<CR>", "r23/O9/F60/S0<CR>", "r23/O10/F60/S0<CR>", "r24/O90/F60/S0<CR>", "r24/O75/F60/S0<CR>", "r25/O675/F60/S0<CR>" };
            //[2,2]
            /*  Frecuencia AC @60 Hz */
            string[] configCalibradoFreq = new string[] { "r22/O1/F10/S0<CR>", "r22/O1/F40/S0<CR>", "r22/O1/F40/S0<CR>", "r22/O1/F300000/S0<CR>" };
            //[2,2]
            /*  Corriente AC @60 Hz */
            string[] configCalibradoAcC = new string[] { "r34/O100/F60/S0<CR>", "r35/O0.9/F60/S0<CR>", "r35/O0.3/F60/S0<CR>", "highcurrent", "r36/O2.7/F60/S0<CR>" };
            //[2,-1,2]
            /*  Corriente DC */
            string[] configCalibradoDcC = new string[] { "r12/O1/S0<CR>", "r13/O9/S0<CR>", "r13/O10/S0<CR>", "r14/O90/S0<CR>", "r14/O100/S0<CR>", "r15/O0.9/S0<CR>", "r15/O0.3/S0<CR>", "highcurrent", "r16/O2.7/S0<CR>" };

            /* Resistencia */
            // CALC:FUNC NULL
            //CALC:STAT ON
            string[] configEquipoRest = new string[] {
                "CONF:RES 100",
                "CONF:RES 1000",
                "CONF:RES 10000",
                "CONF:RES 100000",
                "CONF:RES 1000000",
                "CONF:RES 10000000",
                "CONF:RES 100000000" };
            //CAlC:STAT OFF
            /*  DC V */
            string[] configEquipoDcV = new string[] { "CONF:VOLT:DC 100 mV", "CONF:VOLT:DC 1 V", "CONF:VOLT:DC 10 V", "CONF:VOLT:DC 100 V", "CONF:VOLT:DC 1000 V" };
            /*  AC V  @60 Hz*/
            string[] configEquipoAcV = new string[] { "CONF:VOLT:AC 100 mV", "CONF:VOLT:AC 1 V", "CONF:VOLT:AC 10 V", "CONF:VOLT:AC 100 V", "CONF:VOLT:AC 750 V" };
            /* Fecuencia*/
            string[] configEquipoFreq = new string[] { "CONF:FREQ 40 Hz", "CONF:FREQ 300000 HZ" };
            /*  Corriente AC @60 Hz */
            //Cambiar cables para medir corriente
            string[] configEquipoAcC = new string[] { "CONF:CURR:AC 1 A", "CONF:CURR:AC 3 A" };
            /*  Corriente DC */
            string[] configEquipoDcC = new string[] { "CONF:CURR:DC 10 mA", "CONF:CURR:DC 100 mA", "CONF:CURR:DC 1 A", "CONF:CURR:DC 3 A" };
            #endregion            

            for (int i = 0; i < funcactiva.Length; i++)
            {
                int nrango = rangoxFunc[i];
                int enable = funcactiva[i];

                string[] vmedir_temp = new string[] { "" };
                vmedir_temp = vmedir[i].Split(',');

                string[] tiempovm_temp = new string[] { "" };
                tiempovm_temp = tiempovm[i].Split(',');

                //myGenerador.WriteString("*CLS", true); //Cleaning configuracion
                writeVisa(list_insta[0].device, "*CLS"); // Este device esta en modo generador

                string funcion = nombrefunc[i];

                if (funcion == "Resistencia" && enable == 1)
                {
                    metroTextBox_command.Text += "\r\n" + "Funcion | " + funcion;
                    FunctionReadMultimetro(i, configCalibradorRest, configEquipoRest, nrango, vmedir_temp, tiempovm_temp, modocal[i]);
                    //[numfuncion, comandosGenerador, ComandosLector, rangoxfuncion, numrepeticionesxfuncion,tiempo,mododecalxfuncion]
                    //myEquipo.WriteString("CALC:STAT OFF", true);
                }
                else if (funcion == "DC V" && enable == 1)
                {
                    metroTextBox_command.Text += "\r\n" + "Funcion | " + funcion;
                    FunctionReadMultimetro(i, configCalibradorDcV, configEquipoDcV, nrango, vmedir_temp, tiempovm_temp, modocal[i]);
                }
                else if (funcion == "AC V @60hz" && enable == 1)
                {
                    metroTextBox_command.Text += "\r\n" + "Funcion | " + funcion;
                    FunctionReadMultimetro(i, configCalibradoAcV, configEquipoAcV, nrango, vmedir_temp, tiempovm_temp, modocal[i]);
                }
                else if (funcion == "Frecuencia" && enable == 1)
                {
                    metroTextBox_command.Text += "\r\n" + "Funcion | " + funcion;
                    FunctionReadMultimetro(i, configCalibradoFreq, configEquipoFreq, nrango, vmedir_temp, tiempovm_temp, modocal[i]);
                }
                if (funcion == "AC I @60hz" && enable == 1)
                {
                    metroTextBox_command.Text += "\r\n" + "Pregunta | " + "Mediremos Baja Corriente, Cambiar Cables por favor, cuando estes listo presiona - Yes-?";
                    DialogResult dr = new DialogResult();
                    do
                    {
                        dr = MetroFramework.MetroMessageBox.Show(this,
                      "\n\n Mediremos Baja Corriente, Cambiar Cables por favor, cuando estes listo presiona - Yes-?", "Alerta!",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Question
                        );
                    } while (dr != DialogResult.Yes);

                    if (dr == DialogResult.Yes)
                    {
                        metroTextBox_command.Text += "\r\n" + "Funcion | " + funcion;
                        FunctionReadMultimetro(i, configCalibradoAcC, configEquipoAcC, nrango, vmedir_temp, tiempovm_temp, modocal[i]);
                    }
                    else
                    {
                        //Terminar calibracion                     
                    }
                }
                if (funcion == "DC I" && enable == 1)
                {
                    metroTextBox_command.Text += "\r\n" + "Pregunta | " + "Mediremos Baja Corriente, Cambiar Cables por favor, cuando estes listo presiona- Yes-?";
                    DialogResult dr = new DialogResult();
                    do
                    {
                        dr = MetroFramework.MetroMessageBox.Show(this,
                      "\n\n Mediremos Baja Corriente, Cambiar Cables por favor, cuando estes listo presiona- Yes-?", "Alerta!",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Question
                        );
                    } while (dr != DialogResult.Yes);

                    if (dr == DialogResult.Yes)
                    {
                        metroTextBox_command.Text += "\r\n" + "Funcion | " + funcion;
                        FunctionReadMultimetro(i, configCalibradoDcC, configEquipoDcC, nrango, vmedir_temp, tiempovm_temp, modocal[i]);
                    }
                    else
                    {
                        //Terminar calibracion                     
                    }
                }

            }
        }

        public void FunctionReadMultimetro(int funcion, string[] configCalibrador, string[] configEquipo, int nrango, string[] arrmediciones, string[] arrtiempovm, int modocal)
        {
            //[numfuncion, comandosGenerador, ComandosLector, rangoxfuncion, numrepeticionesxfuncion,tiempo,mododecalxfuncion]

            string[] nombrevrango = new string[] { "100 Ω,1 kΩ,10 kΩ,100 kΩ,1 MΩ,10 MΩ,100 MΩ", "100 mV,1 V,10 V,100 V,1000 V", "100 mV,1 V,10 V,100 V,750 V", "10 Hz - 40 Hz,40 Hz - 300 kHz", "1 A,3 A", "10 mA,100 mA,1 A,3 A" }; //Resistencia, DC V, AC V @60hz, Frecuencia @60hz, Corriente AC @60hz, Corriente DC
            string[] nombrevnominal = new string[] { "10,90,0.1,0.9,1,9,10,90,0.1,0.9,1,9,10,90", "10,-10,90,-90,0.1,-0.1,0.9,-0.9,1,9,10,90,100,900", "10,90,0.1,0.9,1,9,10,90,75,675", "10,40,40,300", "0.1,0.9,0.3,2.7,2.7", "1,9,10,90,0.1,0.9,0.3,2.7,2.7" }; //Resistencia, DC V, AC V @60hz, Frecuencia @60hz, Corriente AC @60hz, Corriente DC
            string[] nombremodocal = new string[] { "Paralelo", "Serie", "Separado" };
            string[] nombrefunc = new string[] { "Resistencia", "DC V", "AC V @60hz", "Frecuencia", "AC I @60hz", "DC I" };

            string[] vrango_temp = new string[] { "" };
            vrango_temp = nombrevrango[funcion].Split(',');

            string[] vnominal_temp = new string[] { "" };
            vnominal_temp = nombrevnominal[funcion].Split(',');

            int n_eqcalibrar = list_instb.Count;

            string retorno = "";
            int pivoteconfigCalib = 0;
            string configCal = "";
            double[] lectura = new double[3];
            bool lecturacorrecta = true;

            for (int a = 0; a < n_eqcalibrar; a++)
            {
                pivoteconfigCalib = 0;
                if (funcion == 0 && modocal == 2)
                {
                    if (n_eqcalibrar > 0) // Equipo por equipo, opcion cuando sean mas de uno
                    {
                        metroTextBox_command.Text += "\r\n" + "Pregunta | " + "Se realizara la calibración del equipo " + list_instb[a].modelo.ToString() + " de la función: " + nombrefunc[funcion] + ", asi que se requiere la conección en " + nombremodocal[modocal] + ". Cuando este listo presiona -Yes-";
                       DialogResult dr = new DialogResult();
                        do
                        {
                            dr = MetroFramework.MetroMessageBox.Show(this,
                          "\n\n Se realizara la calibración del equipo " + list_instb[a].modelo.ToString() + " de la función: " + nombrefunc[funcion] + ", asi que se requiere la conección en " + nombremodocal[modocal] + ". Cuando este listo presiona -Yes-", "Alerta!",
                           MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question
                            );
                        } while (dr != DialogResult.Yes);

                    }
                    /* Escribir la funcion por cada equipo activo, en el archivo correspondiente */
                    //Writefuncion(EquipoName[a].ToString(), nombrefunc[funcion].ToString());

                    Writefuncion(list_instb[a].file.ToString(), nombrefunc[funcion].ToString());                   

                    writeVisa(list_instb[a].device, "*CLS");
                }
                else if (modocal < 2)
                {
                    metroTextBox_command.Text += "\r\n" + "Pregunta | " + "Se realizara la calibración simultanea de la función: " + nombrefunc[funcion] + ", asi que se requiere la conección en " + nombremodocal[modocal] + ". Cuando este listo presiona -Yes-";
                    DialogResult dr = new DialogResult();
                    do
                    {
                        dr = MetroFramework.MetroMessageBox.Show(this,
                      "\n\n Se realizara la calibración simultanea de la función: " + nombrefunc[funcion] + ", asi que se requiere la conección en " + nombremodocal[modocal] + ". Cuando este listo presiona -Yes-", "Alerta!",
                       MessageBoxButtons.YesNo,
                       MessageBoxIcon.Question
                        );
                    } while (dr != DialogResult.Yes);


                    //Aqui se hace la pura configuracion de todos los equipos conectados en cada rango                            
                    for (int b = 0; b < n_eqcalibrar; b++)
                    {
                        /* Escribir las funciones por cada equipo simultaneamente, en el archivo correspondiente */                       
                        Writefuncion(list_instb[b].file.ToString(), nombrefunc[funcion].ToString());
                    }
                }

                for (int i = 0; i < nrango; i++) // Numero de rangos por funcion [7,5,...]
                {
                    int nmediciones = int.Parse(arrmediciones[i]);
                    string[] n_coleccionvpatron = new string[n_eqcalibrar];
                    string[] n_coleccionlecturas = new string[n_eqcalibrar];
                    string coleccionvpatron = "";
                    string coleccionlecturas = "";

                    //myGenerador.WriteString("*CLS", true);                        
                    writeVisa_2(list_insta[0].device, "*RST");
                    writeVisa_2(list_insta[0].device, "*RST");
                    writeVisa_2(list_insta[0].device, "*CLS");
                    //writeVisa(list_insta[0].device, "/S1<CR>");
                    metroTextBox_command.Text += "\r\n" + "Wait | " + "1500";
                   Thread.Sleep(1500);

                    for (int j = 0; j < nmediciones; j++) // numero de mediciones que se van hacer en el rango de la funcion. Ejemplo: Resistencia->100-> #2 -> repeticiones 3 por cada nmediciones
                    {

                        if (funcion == 0 && j == 0)// funcion==0 quiere decir que la funcion es resistencia
                        {
                            //Configuracion para la funcion del equipo por rango 
                            //myGenerador.WriteString("r164/O0/S0<CR>", true);
                            metroTextBox_command.Text += "\r\n" + "Wait | " + "500";
                            Thread.Sleep(500);
                            writeVisa(list_insta[0].device, "r164/O0/S0<CR>");
                            //myEquipo.WriteString("*CLS", true);
                            writeVisa(list_instb[a].device, "*CLS");
                            metroTextBox_command.Text += "\r\n" + "Wait | " + "500";
                            Thread.Sleep(500);
                            writeVisa(list_instb[a].device, "*RST");
                            metroTextBox_command.Text += "\r\n" + "Wait | " + "1000";
                            Thread.Sleep(1000);                            

                            string configtemp = configEquipo[i];
                            writeVisa(list_instb[a].device, configtemp);
                            metroTextBox_command.Text += "\r\n" + "Wait | " + "500";
                            Thread.Sleep(500);
                            //myEquipo.WriteString("CALC:FUNC NULL", true);
                            //myEquipo.WriteString("CALC:STAT ON", true);

                            //writeVisa(list_instb[a].device, "CALC:FUNC NULL");
                            //writeVisa(list_instb[a].device, "CALC:STAT ON");
                            metroTextBox_command.Text += "\r\n" + "Wait | " + "1000";
                            Thread.Sleep(1000);
                                                                                 
                            //metroTextBox_command.Text += "\r\n" + "Equipo a calibrar, NUll configurado"; 
                            Writerango(list_instb[a].file.ToString(), vrango_temp[i].ToString());

                        }
                        else if (funcion > 0 && j == 0)
                        {
                            //Aqui se hace la pura configuracion de todos los equipos conectados en cada rango                            
                            for (int b = 0; b < n_eqcalibrar; b++)
                            {                                                                                                   
                                writeVisa(list_instb[b].device, "*CLS");                                
                                string configtemp = configEquipo[i];
                                writeVisa(list_instb[b].device, configtemp);

                                //Console.WriteLine("Configuración Equipo [" + configtemp + "]");                                    

                                /* Escribir los rangos por cada equipo simultaneamente, en el archivo correspondiente */                                
                                Writerango(list_instb[b].file.ToString(), vrango_temp[i].ToString());
                            }
                        }

                        /* Esta condicion sirve para notificar que hay que cambiar cables*/
                        if (funcion > 3 && configCalibrador[pivoteconfigCalib] == "highcurrent")
                        {
                            //myGenerador.WriteString("/S1<CR>", true); // Poner en Standby
                            writeVisa(list_insta[0].device, "*CLS");
                            writeVisa_2(list_insta[0].device, "/S1<CR>");

                            metroTextBox_command.Text += "\r\n" + "Pregunta | " + "Mediremos Alta Corriente, favor de cambiar los cables a las terminales correspondientes. Cuando estes listo listo presiona -Yes-?";
                            DialogResult dr = new DialogResult();
                            do
                            {
                                dr = MetroFramework.MetroMessageBox.Show(this,
                              "\n\n Mediremos Alta Corriente, favor de cambiar los cables a las terminales correspondientes. Cuando estes listo listo presiona -Yes-?", "Alerta!",
                               MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question
                                );
                            } while (dr != DialogResult.Yes);

                            // debug si estra a esta opción
                            if (dr == DialogResult.Yes)
                            {
                                pivoteconfigCalib = pivoteconfigCalib + 1;
                            }

                        }

                        //Configuracion para el Calibrador
                        configCal = configCalibrador[pivoteconfigCalib];
                        //myGenerador.WriteString(configCal, true);
                        writeVisa_2(list_insta[0].device, configCal);
                        writeVisa_2(list_insta[0].device, "*WAI");
                        metroTextBox_command.Text += "\r\n" + "Wait | " + "1500";
                        Thread.Sleep(1500);                        
                        int delay = int.Parse(arrtiempovm[i]);

                        Thread.Sleep(delay);

                        if (funcion == 0)
                        {
                            coleccionvpatron = coleccionvpatron + vnominal_temp[pivoteconfigCalib] + ",";

                            for (int k = 0; k < 3; k++)
                            {
                                //Reading...                             
                                writeVisa(list_instb[a].device, "READ?");
                                retorno = readVisa(list_instb[a].device);
                                lectura[k] = Convert.ToDouble(retorno);
                                metroTextBox_lecturas.Text += "\r\n" + lectura[k].ToString();
                                metroTextBox_command.Text += "\r\n" + "Wait | " + "500";
                                Thread.Sleep(500);
                            }

                            if (lecturacorrecta == true)
                            {
                                for (int c = 0; c < lectura.Length; c++)
                                {
                                    coleccionlecturas = coleccionlecturas + lectura[c].ToString() + ",";
                                }
                            }

                        }
                        else
                        {
                            //Aqui se hace la lectura de los equipos conectados en cada rango                            
                            for (int b = 0; b < n_eqcalibrar; b++)
                            {
                                n_coleccionvpatron[b] = n_coleccionvpatron[b] + vnominal_temp[pivoteconfigCalib] + ",";

                                lecturacorrecta = true;
                                //myEquipo.IO = (IMessage)rm.Open(EquipoAddrs[b], AccessMode.NO_LOCK, 2000, ""); //Abra un controlador para el DMM con un tiempo de espera de 2 segundos

                                Thread.Sleep(1000);
                                for (int k = 0; k < 3; k++)
                                {
                                    //Reading...
                                    //retorno = Read(myEquipo);
                                    // REvisar error

                                    writeVisa(list_instb[b].device, "READ?");
                                    retorno = readVisa(list_instb[b].device);
                                    lectura[k] = Convert.ToDouble(retorno);
                                    metroTextBox_lecturas.Text += "\r\n" + lectura[k].ToString();
                                    metroTextBox_command.Text += "\r\n" + "Wait | " + "500";
                                    Thread.Sleep(500);
                                }

                                if (lecturacorrecta == true)
                                {
                                    for (int c = 0; c < lectura.Length; c++)
                                    {
                                        n_coleccionlecturas[b] = n_coleccionlecturas[b] + lectura[c].ToString() + ",";
                                    }

                                }
                            }
                            a = n_eqcalibrar;
                        }
                        pivoteconfigCalib = pivoteconfigCalib + 1;
                    }

                    if (funcion == 0)
                    {
                        //excel.Open(EquipoName[a].ToString());                           
                        metroTextBox_command.Text += "\r\n" + "Guardando lectura ...";
                        Writeread(list_instb[a].file.ToString(), coleccionvpatron, coleccionlecturas);
                        // Writeread(EquipoName[a].ToString(), coleccionvpatron, coleccionlecturas);
                    }
                    else
                    {
                        for (int b = 0; b < n_eqcalibrar; b++)
                        {
                            //Writeread(EquipoName[b].ToString(), n_coleccionvpatron[b], n_coleccionlecturas[b]);                                                                                    
                            metroTextBox_command.Text += "\r\n" + "Guardando lectura ...";
                            Writeread(list_instb[b].file.ToString(), n_coleccionvpatron[b], n_coleccionlecturas[b]);
                        }
                    }

                }

                for (int c = 0; c < n_eqcalibrar; c++)
                {
                    //myEquipo.IO = (IMessage)rm.Open(EquipoAddrs[c], AccessMode.NO_LOCK, 2000, "");
                    //myEquipo.WriteString("*RST", true);
                    writeVisa(list_instb[c].device, "*RST");

                }
                //myGenerador.WriteString("*RST", true);
                //writeVisa(list_insta[0].device, "*RST");
                metroTextBox_command.Text += "\r\n" + "Wait | " + "2000";
                Thread.Sleep(2000);
            }
            metroTextBox_command.Text += "\r\n" + "Wait | " + "3000";
            Thread.Sleep(3000);
            // StandBy
            // myCalibrador.WriteString("/S1<CR>", true); // Poner en Standby   
            //writeVisa(list_insta[0].device, "/S1<CR>");
            writeVisa_2(list_insta[0].device, "*RST");
        }


        #endregion


        #region Funciones de Write, Read Ivi.Visa

        public IMessageBasedSession openVisa(string resource)
        {
            var mydevice = GlobalResourceManager.Open(resource) as IMessageBasedSession;
            return mydevice;
        }

        public void closeVisa(IMessageBasedSession mydevice)
        {
            mydevice.Dispose(); // write to instrument
        }

        public void writeVisa(IMessageBasedSession mydevice, string command)
        {
            //mydevice.TimeoutMilliseconds = 1000;
            mydevice.RawIO.Write(command); // write to instrument 
            metroTextBox_command.Text += "\r\n" + mydevice.ResourceName + "|"+ command;
        }

        public void writeVisa_2(IMessageBasedSession mydevice, string command)
        {
            // mydevice.TimeoutMilliseconds = 2000;
            mydevice.Clear();
            mydevice.RawIO.Write(command + "\n");
            mydevice.FormattedIO.Printf(command + "\n");
            metroTextBox_command.Text += "\r\n" + mydevice.ResourceName + "|" + command;
            //mydevice.RawIO.Write(command); // write to instrument              
        }

        public string readVisa(IMessageBasedSession mydevice)
        {
            string result = "0";
            //mydevice.TimeoutMilliseconds = 1000;
            result = mydevice.RawIO.ReadString();
            Console.WriteLine("resultado:"+ result);            
            return result; // read from instrument
        }

        #endregion

        #region Modulos para creacion y escritura de las lecturas de acuerdo al equipo, funcion, rango, vpatron

        public static void Createfile(string nombre)
        {
            excel.CreateExel(nombre);
        }

        public static void Writefuncion(string nombreFile, string funcion)
        {
            // Excel excel = new Excel();
            excel.Open(nombreFile);
            int nfila = excel.ScanCell();
            nfila = (nfila == 1) ? 1 : (nfila + 1);
            excel.WriteFuncion(nfila, funcion);
            excel.Save();
            //nfila++;
            //excel.WriteRango(nfila, "Rango 0");
            //excel.Save();
            //excel.WriteValorLectura(nfila, "Valor patron1, valor patron2", "lectura1, lectura2,lectura3,lectura4");
            //excel.Save();
            //excel.SaveAs(@"C:\Users\840HP\Downloads\Csharp-Example\VS2005\BasicFunctions\bin\Debug\prueba2.xlsx");
            excel.Close();
        }

        public static void Writerango(string nombreFile, string rango)
        {
            //Excel excel = new Excel();
            excel.Open(nombreFile);
            int nfila = excel.ScanCell();
            nfila = (nfila == 1) ? 2 : (nfila + 1);
            excel.WriteRango(nfila, rango);
            excel.Save();
            excel.Close();
        }

        public static void Writeread(string nombreFile, string vpatron, string lecturas)
        {
            //Excel excel = new Excel();
            excel.Open(nombreFile);
            int nfila = excel.ScanCell();
            nfila = nfila + 1;
            int sizevp = vpatron.Length;
            vpatron = vpatron.Substring(0, sizevp - 1);
            int sizel = lecturas.Length;
            lecturas = lecturas.Substring(0, sizel - 1);
            excel.WriteValorLectura(nfila, vpatron, lecturas);
            excel.Save();
            excel.Close();
        }

        #endregion

        private void metroButton_cancelar_Click(object sender, EventArgs e)
        {
            this.Clear();
        }

        private void metroButton_refresh_Click(object sender, EventArgs e)
        {
            metroCombodevices.Items.Clear();
            resorce();
        }

        public void Clear()
        {
            metroCombodevices.Text = "-- Seleccionar --";
            metroLabel_idn.Text = "";
            metroRadioButton_patron.Checked = false;
            metroRadioButton_bc.Checked = false;

            metroRadioButton_generador.Checked = false;
            metroRadioButton_lector.Checked = false;

            metroTextBox_informe.Enabled = false;
            metroTextBox_informe.Text = "";
        }

        private void metroButton_limpiarequipos_Click(object sender, EventArgs e)
        {
            //if (list_insta.Count > 0)
            //{
            //    list_insta.Clear();
            //    showlista();
            //}
            //else if (list_instb.Count > 0)
            //{
            //    list_instb.Clear();
            //    showlistb();                
            //}
            //else {
            //    MetroMessageBox.Show(this, "No hay equipos que eliminar!", "Mensaje de notificación", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
        }
    }
}


