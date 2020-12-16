using Microsoft.SolverFoundation.Services;
using OSIsoft.AF;
using OSIsoft.AF.Asset;
using OSIsoft.AF.Search;
using OSIsoft.AF.Time;
using OSIsoft.AF.UnitsOfMeasure;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace NivelLLenadoBolasSag
{
    partial class NivelBolasSag : ServiceBase
    {

        bool blBandera = false;
        public NivelBolasSag()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // TODO: agregar código aquí para iniciar el servicio.
            stLapso.Start();
        }

        protected override void OnStop()
        {
            // TODO: agregar código aquí para realizar cualquier anulación necesaria para detener el servicio.
            stLapso.Stop();
        }

        private void stLapso_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (blBandera) return;

            try
            {
                blBandera = true;

                EventLog.WriteEntry("Se inicio proceso de Calculo Nivel LLenado de Bolas SAG", EventLogEntryType.Information);
                CalculosLLenadoBolasSAG();
            }
            catch (Exception ex)
            {
                //oLog.Add(ex.Message);
                EventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
            }

            blBandera = false;

        }

        private void CalculosLLenadoBolasSAG()
        {
            string path = ConfigurationManager.AppSettings["path"];
            Log oLog = new Log(path);

            string servidor = ConfigurationManager.AppSettings["ServidorAF"];
            string usuario = ConfigurationManager.AppSettings["Usuario"];
            string password = ConfigurationManager.AppSettings["Password"];
            string dominio = ConfigurationManager.AppSettings["Dominio"];
            string bd = ConfigurationManager.AppSettings["BD"];
            string modelo = ConfigurationManager.AppSettings["Modelo"];
            string servidorpi = ConfigurationManager.AppSettings["ServidorPI"];

            AFElement model = new AFElement();
            PISystems AF = new PISystems();
            PISystem AFSrv = AF[servidor];
            string AFConnectionStatus = string.Empty;
            string AFConnectionStatusError = string.Empty;

            try
            {
                NetworkCredential credential = new NetworkCredential(usuario, password, dominio);
                AFSrv.Connect(credential);
                AFConnectionStatus = "Good";
                AFConnectionStatusError = " Conexión Exitosa a Servidor AF " + AFSrv.Name;
                oLog.Add(AFConnectionStatusError);

            }
            catch (Exception ex)
            {
                // Expected exception since credential needs a valid user name and password.
                AFConnectionStatus = "Failed";
                AFConnectionStatusError = " Servidor AF no existe o Revisar Credenciales en Archivo .ini" + " MsgError:" + ex.Message;
                oLog.Add(AFConnectionStatusError);
                //oLog.Add("Fin de Ciclo");
                AFSrv.Disconnect();
                AFSrv.Dispose();
                //Environment.Exit(0);
            }

            //Elegir Base de Datos AF
            string AFConnectionStatusBD = string.Empty;
            string AFConnectionStatusBDError = string.Empty;
            string templateName = "MolinoSag";
            if (AFConnectionStatus == "Good")
            {
                AFDatabases AFSrvBDs = AFSrv.Databases;
                AFDatabase AFSrvBD = AFSrvBDs[bd];

                oLog.Add("---Inicio Proceso---");

                using (AFElementSearch elementQuery = new AFElementSearch(AFSrvBD, "TemplateSearch", string.Format("template:\"{0}\"", templateName)))
                {
                    //elementQuery.CacheTimeout = TimeSpan.FromMinutes(5);
                    
                    foreach (AFElement element in elementQuery.FindElements())
                    {

                        try
                        {
                            AFTimeRange timeRange = new AFTimeRange();
                            
                            timeRange.StartTime = DateTime.Now.AddDays(-1);
                            timeRange.EndTime = DateTime.Now;

                            //Traer Data
                            AFValues M_Potencia = element.Attributes["Potencia"].GetValues(timeRange, -720, null);
                            AFValues M_Velocidad = element.Attributes["Velocidad"].GetValues(timeRange, -720, null);
                            AFValues M_presion = element.Attributes["Presion"].GetValues(timeRange, -720, null);
                            AFValues M_TPH = element.Attributes["Tph"].GetValues(timeRange, -720, null);
                            AFValues M_TPH_Fresco = element.Attributes["Tph Fresco"].GetValues(timeRange, -720, null);
                            AFValues M_CP = element.Attributes["%Cp"].GetValues(timeRange, -720, null);

                            //Filtrar Data
                            double potencia1;
                            double presion1;
                            double velocidad1;
                            double tph_1;
                            double tph_f1;
                            double CP_1;
                            int Valores_Bolas = 0;
                            //int fin = 0;

                            double[] M_Potencia_Filt = new double[725];
                            double[] M_presion_Filt = new double[725];
                            double[] M_Velocidad_Filt = new double[725];
                            double[] M_TPH_Filt = new double[725];
                            double[] M_TPH_Fresco_Filt = new double[725];
                            double[] M_CP_Filt = new double[725];

                            double M_Potencia_Min = element.Attributes["Potencia_Min"].GetValue().ValueAsDouble();
                            double M_presion_Min = element.Attributes["Presion_Min"].GetValue().ValueAsDouble();
                            double M_Velocidad_Min = element.Attributes["Velocidad_Min"].GetValue().ValueAsDouble();
                            double M_TPH_Min = element.Attributes["Tph_Min"].GetValue().ValueAsDouble();
                            double M_TPH_Fresco_Min = element.Attributes["Tph Fresco_Min"].GetValue().ValueAsDouble();
                            double M_CP_Min = element.Attributes["Cp_Min"].GetValue().ValueAsDouble();

                            for (int x=0; x <= 719; x++)
                            {
                                try
                                {
                                    potencia1 = M_Potencia[x].ValueAsDouble();
                                }
                                catch
                                {
                                    potencia1 = 0;
                                }
                                try
                                {
                                    presion1 = M_presion[x].ValueAsDouble();
                                }
                                catch
                                {
                                    presion1 = 0;
                                }
                                try
                                {
                                    velocidad1 = M_Velocidad[x].ValueAsDouble();
                                }
                                catch
                                {
                                    velocidad1 = 0;
                                }
                                try
                                {
                                    tph_1 = M_TPH[x].ValueAsDouble();
                                }
                                catch
                                {
                                    tph_1 = 0;
                                }
                                try
                                {
                                    tph_f1 = M_TPH_Fresco[x].ValueAsDouble();
                                }
                                catch
                                {
                                    tph_f1 = 0;
                                }
                                try
                                {
                                    CP_1 = M_CP[x].ValueAsDouble();
                                }
                                catch
                                {
                                    CP_1 = 0;
                                }

                                if(potencia1 > M_Potencia_Min && presion1 > M_presion_Min && velocidad1 > M_Velocidad_Min && tph_1 > M_TPH_Min && tph_f1 > M_TPH_Fresco_Min && CP_1 > M_CP_Min)
                                {
                                    Valores_Bolas = Valores_Bolas + 1;
                                    M_Potencia_Filt[Valores_Bolas] = M_Potencia[x].ValueAsDouble();
                                    M_presion_Filt[Valores_Bolas] =  M_presion[x].ValueAsDouble();
                                    M_Velocidad_Filt[Valores_Bolas] = M_Velocidad[x].ValueAsDouble();
                                    M_TPH_Filt[Valores_Bolas] = M_TPH[x].ValueAsDouble();
                                    M_TPH_Fresco_Filt[Valores_Bolas] = M_TPH_Fresco[x].ValueAsDouble();
                                    M_CP_Filt[Valores_Bolas] = M_CP[x].ValueAsDouble();
                                }

                            }

                            
                            //Calcular Intervalos
                            //double potencia;
                            //int X, y;
                            double Valor_Max=0, Valor_Max1=0;


                            //AFValue Inter_Max = element.Attributes["Inter_Max"].GetValue();
                            //AFValue Inter_Min = element.Attributes["Inter_Min"].GetValue();

                            //int num;
                            //double N1, N2, N3, N4, N5, N6, N7, N8, N9, N10, N11;
                            //double M1, M2, M3, M4, M5, M6, M7, M8, M9, M10, M11;

                            //try
                            //{
                            //  //N1 = 0;
                            //}
                            //catch
                            //{
                            //   //N1 = 0;
                            //}

                            //Este código al parecer no sirve para nada
                            //String Variable_Evalua = element.Attributes["VariablePivot"].GetValue().ToString();
                            //if(Variable_Evalua == "Velocidad")
                            //{
                            //    num = 4;
                            //}

                            double sum00=0, sum001=0, sum002=0, sum003=0, sum004=0, sum005=0, f003=0, nn=0;
                            double sum0=0, sum01=0, sum02 = 0, sum03=0, sum04 = 0, sum05 = 0, f03 = 0, n = 0;
                            double sum1=0, sum11 = 0, sum12 = 0, sum13 = 0, sum14 = 0, sum15 = 0, f13 = 0, a = 0;
                            double sum2=0, sum21 = 0, sum22 = 0, sum23 = 0, sum24 = 0, sum25 = 0, f23 = 0, b = 0;
                            double sum3=0, sum31 = 0, sum32 = 0, sum33 = 0, sum34 = 0, sum35 = 0, f33 = 0, c = 0;
                            double sum4=0, sum41 = 0, sum42 = 0, sum43 = 0, sum44 = 0, sum45 = 0, f43 = 0, d = 0;
                            double sum5=0, sum51 = 0, sum52 = 0, sum53 = 0, sum54 = 0, sum55 = 0, f53 = 0, e = 0;
                            double sum6=0, sum61 = 0, sum62 = 0, sum63 = 0, sum64 = 0, sum65 = 0, f63 = 0, f = 0;
                            double sum7=0, sum71 = 0, sum72 = 0, sum73 = 0, sum74 = 0, sum75 = 0, f73 = 0, g = 0;
                            double sum8=0, sum81 = 0, sum82 = 0, sum83 = 0, sum84 = 0, sum85 = 0, f83 = 0, h = 0;
                            double sum9=0, sum91 = 0, sum92 = 0, sum93 = 0, sum94 = 0, sum95 = 0, f93 = 0, i = 0;

                            for(int xx=1; xx <= Valores_Bolas; xx++)
                            {
                                if (M_Velocidad_Filt[xx] > 6 && M_Velocidad_Filt[xx] < 6.5)
                                {
                                    sum00 +=  M_presion_Filt[xx];
                                    sum001 +=  M_Velocidad_Filt[xx];
                                    sum002 +=  M_Potencia_Filt[xx];
                                    sum003 +=  M_TPH_Filt[xx];
                                    f003 += 1;
                                    sum004 += M_TPH_Fresco_Filt[xx];
                                    sum005 += + M_CP_Filt[xx];
                                    nn += 1;
                                }

                               if(M_Velocidad_Filt[xx] > 6.5 && M_Velocidad_Filt[xx] < 7)
                                {
                                    sum0 = sum0 + M_presion_Filt[xx];
                                    sum01 = sum01 + M_Velocidad_Filt[xx];
                                    sum02 = sum02 + M_Potencia_Filt[xx];
                                    sum03 = sum03 + M_TPH_Filt[xx];
                                    f03 = f03 + 1;
                                    sum04 = sum04 + M_TPH_Fresco_Filt[xx];
                                    sum05 = sum05 + M_CP_Filt[xx];
                                    n = n + 1;
                                }

                                if(M_Velocidad_Filt[xx] > 7 && M_Velocidad_Filt[xx] < 7.5)
                                {
                                    sum1 = sum1 + M_presion_Filt[xx];
                                    sum11 = sum11 + M_Velocidad_Filt[xx];
                                    sum12 = sum12 + M_Potencia_Filt[xx];
                                    sum13 = sum13 + M_TPH_Filt[xx];
                                    f13 = f13 + 1;
                                    sum14 = sum14 + M_TPH_Fresco_Filt[xx];
                                    sum15 = sum15 + M_CP_Filt[xx];
                                    a = a + 1;
                                }

                                if(M_Velocidad_Filt[xx] > 7.5 && M_Velocidad_Filt[xx] < 8)
                                {
                                    sum2 = sum2 + M_presion_Filt[xx];
                                    sum21 = sum21 + M_Velocidad_Filt[xx];
                                    sum22 = sum22 + M_Potencia_Filt[xx];
                                    sum23 = sum23 + M_TPH_Filt[xx];
                                    f23 = f23 + 1;
                                    sum24 = sum24 + M_TPH_Fresco_Filt[xx];
                                    sum25 = sum25 + M_CP_Filt[xx];
                                    b = b + 1;
                                }

                                if(M_Velocidad_Filt[xx] > 8 && M_Velocidad_Filt[xx] < 8.5)
                                {
                                    sum3 = sum3 + M_presion_Filt[xx];
                                    sum31 = sum31 + M_Velocidad_Filt[xx];
                                    sum32 = sum32 + M_Potencia_Filt[xx];
                                    sum33 = sum33 + M_TPH_Filt[xx];
                                    f33 = f33 + 1;
                                    sum34 = sum34 + M_TPH_Fresco_Filt[xx];
                                    sum35 = sum35 + M_CP_Filt[xx];
                                    c = c + 1;
                                }

                                if(M_Velocidad_Filt[xx] > 8.5 && M_Velocidad_Filt[xx] < 9)
                                {
                                    sum4 = sum4 + M_presion_Filt[xx];
                                    sum41 = sum41 + M_Velocidad_Filt[xx];
                                    sum42 = sum42 + M_Potencia_Filt[xx];
                                    sum43 = sum43 + M_TPH_Filt[xx];
                                    f43 = f43 + 1;
                                    sum44 = sum44 + M_TPH_Fresco_Filt[xx];
                                    sum45 = sum45 + M_CP_Filt[xx];
                                    d = d + 1;
                                }


                                if(M_Velocidad_Filt[xx] > 9 && M_Velocidad_Filt[xx] < 9.5)
                                {
                                    sum5 = sum5 + M_presion_Filt[xx];
                                    sum51 = sum51 + M_Velocidad_Filt[xx];
                                    sum52 = sum52 + M_Potencia_Filt[xx];
                                    sum53 = sum53 + M_TPH_Filt[xx];
                                    f53 = f53 + 1;
                                    sum54 = sum54 + M_TPH_Fresco_Filt[xx];
                                    sum55 = sum55 + M_CP_Filt[xx];
                                    e = e + 1;
                                }


                                if(M_Velocidad_Filt[xx] > 9.5 && M_Velocidad_Filt[xx] < 10)
                                {
                                    sum6 = sum6 + M_presion_Filt[xx];
                                    sum61 = sum61 + M_Velocidad_Filt[xx];
                                    sum62 = sum62 + M_Potencia_Filt[xx];
                                    sum63 = sum63 + M_TPH_Filt[xx];
                                    f63 = f63 + 1;
                                    sum64 = sum64 + M_TPH_Fresco_Filt[xx];
                                    sum65 = sum65 + M_CP_Filt[xx];
                                    f = f + 1;
                                }

                                if(M_Velocidad_Filt[xx] > 10 && M_Velocidad_Filt[xx] < 10.5)
                                {
                                    sum7 = sum7 + M_presion_Filt[xx];
                                    sum71 = sum71 + M_Velocidad_Filt[xx];
                                    sum72 = sum72 + M_Potencia_Filt[xx];
                                    sum73 = sum73 + M_TPH_Filt[xx];
                                    f73 = f73 + 1;
                                    sum74 = sum74 + M_TPH_Fresco_Filt[xx];
                                    sum75 = sum75 + M_CP_Filt[xx];
                                    g = g + 1;
                                }

                                if(M_Velocidad_Filt[xx] > 10.5 && M_Velocidad_Filt[xx] < 11)
                                {
                                    sum8 = sum8 + M_presion_Filt[xx];
                                    sum81 = sum81 + M_Velocidad_Filt[xx];
                                    sum82 = sum82 + M_Potencia_Filt[xx];
                                    sum83 = sum83 + M_TPH_Filt[xx];
                                    f83 = f83 + 1;
                                    sum84 = sum84 + M_TPH_Fresco_Filt[xx];
                                    sum85 = sum85 + M_CP_Filt[xx];
                                    h = h + 1;
                                }

                                if(M_Velocidad_Filt[xx] > 11 && M_Velocidad_Filt[xx] < 11.5)
                                {
                                    sum9 = sum9 + M_presion_Filt[xx];
                                    sum91 = sum91 + M_Velocidad_Filt[xx];
                                    sum92 = sum92 + M_Potencia_Filt[xx];
                                    sum93 = sum93 + M_TPH_Filt[xx];
                                    f93 = f93 + 1;
                                    sum94 = sum94 + M_TPH_Fresco_Filt[xx];
                                    sum95 = sum95 + M_CP_Filt[xx];
                                    i = i + 1;
                                }

                            } //Fin for

                            double[,] Intervalos = new double[11, 13];

                            if(nn > 0)
                            {
                                Intervalos[0, 0] = sum00 / nn;
                                Intervalos[0, 1] = sum001 / nn;
                                Intervalos[0, 2] = sum002 / nn;
                                if (f003 > 0) 
                                {
                                    Intervalos[0, 3] = sum003 / f003;
                                }
                                Intervalos[0, 4] = sum004 / nn;
                                Intervalos[0, 5] = sum005 / nn;
                                Intervalos[0, 6] = nn;
                            }

                            if (n > 0) 
                            {
                                Intervalos[1, 0] = sum0 / n;
                                Intervalos[1, 1] = sum01 / n;
                                Intervalos[1, 2] = sum02 / n;
                                if (f03 > 0)
                                {
                                    Intervalos[1, 3] = sum03 / f03;
                                }
                                Intervalos[1, 4] = sum04 / n;
                                Intervalos[1, 5] = sum05 / n;
                                Intervalos[1, 6] = n;
                            }

                            if (a > 0) 
                            {
                                Intervalos[2, 0] = sum1 / a;
                                Intervalos[2, 1] = sum11 / a;
                                Intervalos[2, 2] = sum12 / a;
                                if(f13 > 0)
                                {
                                    Intervalos[2, 3] = sum13 / f13;
                                }
                                Intervalos[2, 4] = sum14 / a;
                                Intervalos[2, 5] = sum15 / a;
                                Intervalos[2, 6] = a;
                            }

                            if (b > 0)
                            {
                                Intervalos[3, 0] = sum2 / b;
                                Intervalos[3, 1] = sum21 / b;
                                Intervalos[3, 2] = sum22 / b;
                                if(f23 > 0)
                                {
                                    Intervalos[3, 3] = sum23 / f23;
                                }
                                Intervalos[3, 4] = sum24 / b;
                                Intervalos[3, 5] = sum25 / b;
                                Intervalos[3, 6] = b;
                            }

                            if (c > 0)
                            {
                                Intervalos[4, 0] = sum3 / c;
                                Intervalos[4, 1] = sum31 / c;
                                Intervalos[4, 2] = sum32 / c;
                                if (f33 > 0)
                                {
                                    Intervalos[4, 3] = sum33 / f33;
                                }
                                Intervalos[4, 4] = sum34 / c;
                                Intervalos[4, 5] = sum35 / c;
                                Intervalos[4, 6] = c;
                            }

                            if (d > 0)
                            {
                                Intervalos[5, 0] = sum4 / d;
                                Intervalos[5, 1] = sum41 / d;
                                Intervalos[5, 2] = sum42 / d;
                                if (f43 > 0)
                                {
                                    Intervalos[5, 3] = sum43 / f43;
                                }
                                Intervalos[5, 4] = sum44 / d;
                                Intervalos[5, 5] = sum45 / d;
                                Intervalos[5, 6] = d;
                            }

                            if (e > 0)
                            {
                                Intervalos[6, 0] = sum5 / e;
                                Intervalos[6, 1] = sum51 / e;
                                Intervalos[6, 2] = sum52 / e;
                                if (f53 > 0) 
                                {
                                    Intervalos[6, 3] = sum53 / f53;
                                }
                                Intervalos[6, 4] = sum54 / e;
                                Intervalos[6, 5] = sum55 / e;
                                Intervalos[6, 6] = e;
                            }

                            if (f > 0)
                            {
                                Intervalos[7, 0] = sum6 / f;
                                Intervalos[7, 1] = sum61 / f;
                                Intervalos[7, 2] = sum62 / f;
                                if (f63 > 0)
                                {
                                    Intervalos[7, 3] = sum63 / f63;
                                }
                                Intervalos[7, 4] = sum64 / f;
                                Intervalos[7, 5] = sum65 / f;
                                Intervalos[7, 6] = f;
                            }

                            if (g > 0)
                            {
                                Intervalos[8, 0] = sum7 / g;
                                Intervalos[8, 1] = sum71 / g;
                                Intervalos[8, 2] = sum72 / g;
                                if (f73 > 0)
                                {
                                    Intervalos[8, 3] = sum73 / f73;
                                }
                                Intervalos[8, 4] = sum74 / g;
                                Intervalos[8, 5] = sum75 / g;
                                Intervalos[8, 6] = g;
                            }

                            if (h > 0)
                            {
                                Intervalos[9, 0] = sum8 / h;
                                Intervalos[9, 1] = sum81 / h;
                                Intervalos[9, 2] = sum82 / h;
                                if (f83 > 0)
                                {
                                    Intervalos[9, 3] = sum83 / f83;
                                }
                                Intervalos[9, 4] = sum84 / h;
                                Intervalos[9, 5] = sum85 / h;
                                Intervalos[9, 6] = h;
                            }

                            if (i > 0)
                            {
                                Intervalos[10, 0] = sum9 / i;
                                Intervalos[10, 1] = sum91 / i;
                                Intervalos[10, 2] = sum92 / i;
                                if (f93 > 0)
                                {
                                    Intervalos[10, 3] = sum93 / f93;
                                }
                                Intervalos[10, 4] = sum94 / i;
                                Intervalos[10, 5] = sum95 / i;
                                Intervalos[10, 6] = i;
                            }

                            int Select_intervalo = 0;
                            int Select_intervalo1 = 0;
                            for(int t=1; t<=10; t++)
                            {
                                if(Intervalos[t, 6] > Intervalos[t - 1, 6])
                                {
                                    Select_intervalo = t;
                                }
                            }

                            for(int t=1; t<=10; t++)
                            {
                                Valor_Max = Intervalos[t, 6];
                                if (Select_intervalo != t)
                                {
                                    if(Valor_Max > Valor_Max1)
                                    {
                                        Select_intervalo1 = t;
                                    }
                                }
                            }


                            //Estimacion JB
                            //bool estado;
                            int ejecuciones=0;
                            int T;
                            //AFValue valor;
                            int t1, t2;
                            double Rpm=0, Chargue=0, Potencia=0, Solido=0;
                            double mill_speed=0;
                            double jb=0, jb2=0, jb1=0;

                            double diametro = element.Attributes["Diametro"].GetValue().ValueAsDouble();
                            double largo = element.Attributes["Largo"].GetValue().ValueAsDouble();
                            double d_bolas = element.Attributes["Densidad_Bolas"].GetValue().ValueAsDouble();
                            double d_mineral = element.Attributes["Densidad_Mineral"].GetValue().ValueAsDouble();
                            double angulo = element.Attributes["Angulo"].GetValue().ValueAsDouble();
                            double perdida = element.Attributes["Perdida_Energia"].GetValue().ValueAsDouble();
                            double inter = element.Attributes["LLenado_Intersticios"].GetValue().ValueAsDouble();
                            double inferido = element.Attributes["Jc Inferido"].GetValue().ValueAsDouble();

                            T = Select_intervalo;
                            if (Intervalos[T, 0] > 0)
                            {
                                for(int jc=30; jc<=30; jc++)
                                {
                                    Rpm = Intervalos[T, 1];
                                    Chargue = jc;
                                    if (Intervalos[T, 2] < 1000)
                                    {
                                        Potencia = Intervalos[T, 2] * 1000;
                                    }
                                    else
                                    {
                                        Potencia = Intervalos[T, 2];
                                    }

                                    Solido = Intervalos[T, 5];

                                   
                                    try
                                    {
                                        mill_speed = (Solver_1(Rpm, diametro));
                                    }
                                    catch (Exception ex)
                                    {
                                        oLog.Add(ex.Message);
                                    }

                                    try
                                    {

                                        jb = (Solver_Potencia(Potencia, diametro, largo, mill_speed, d_bolas, Solido, d_mineral, angulo, perdida, inter, inferido));
                                    }
                                    catch (Exception ex)
                                    {
                                        oLog.Add(ex.Message);
                                    }

                                    ejecuciones = ejecuciones + 1;

                                    jb1 = jb;

                                }
                            }

                            T = Select_intervalo1;
                            if(Intervalos[T, 0] > 0)
                            {
                                Rpm = Intervalos[T, 1];
                                //Chargue = jc;
                                if (Intervalos[T, 2] < 1000)
                                {
                                    Potencia = Intervalos[T, 2] * 1000;
                                }
                                else
                                {
                                    Potencia = Intervalos[T, 2];
                                }

                                Solido = Intervalos[T, 5];
                                try
                                {
                                    mill_speed = (Solver_1(Rpm, diametro));
                                }
                                catch (Exception ex)
                                {
                                    oLog.Add(ex.Message);
                                }

                                try
                                {
                                    jb = (Solver_Potencia(Potencia, diametro, largo, mill_speed, d_bolas, Solido, d_mineral, angulo, perdida, inter, inferido));
                                }
                                catch (Exception ex)
                                {
                                    oLog.Add(ex.Message);
                                }

                                ejecuciones = ejecuciones + 1;

                                jb2 = jb;

                            }

                            t1 = Select_intervalo;
                            t2 = Select_intervalo1;

                            jb = (jb1 * Intervalos[t1, 6] + jb2 * Intervalos[t2, 6]) / (Intervalos[t1, 6] + Intervalos[t2, 6]);

                            element.Attributes["Jb"].SetValue(new AFValue(jb));

                            oLog.Add("Molino: " + element.Name + " Mill_Speed: " + mill_speed + " Jb: " + jb);
                           
                        }
                        catch (Exception ex)
                        {
                            oLog.Add(ex.Message);
                            //AFSrv.Disconnect();
                           //AFSrv.Dispose();
                        }

                    }

                    oLog.Add("---Fin Proceso---");
                    //AFSrv.Disconnect();
                    //AFSrv.Dispose();
                }

            }

        }


        public static double Solver_1(double rpm, double diametro)
        {
            SolverContext context = SolverContext.GetContext();
            context.ClearModel();
            Model model = context.CreateModel();

            Decision x = new Decision(Domain.RealNonnegative, "Mill_Speed");

            model.AddDecisions(x);
            model.AddConstraints("one", rpm == (76.6 / Math.Pow(diametro, 0.5)) * (x / (double)100));

            SimplexDirective simplex = new SimplexDirective();

            // Solve the problem
            context.Solve(simplex);

            return x.GetDouble();

        }

        public static double Solver_Potencia(double potencia, double diametro, double largo, double millSpeed, double densidad_bolas, double solidos, double densidad_mineral, double angulo, double losses, double inter, double jci)
        {

            SolverContext context1 = SolverContext.GetContext();
            context1.ClearModel();
            Model model1 = context1.CreateModel();

            Decision y = new Decision(Domain.RealNonnegative, "Balls_Filling");

            model1.AddDecisions(y);
                                        //potencia = (0.238 * 38 ^ 3.5 * (22 / 38) * (Mill_Speed / 100) * ((((1 - 0.4) * 7.75 * (y / 100) * Math.PI * (38 * 0.305) ^ 2 *                                                                          (22 * 0.305) / 4) + ((1 - 0.4) * 2.7 * (jc / 100 - y / 100) *                                                                                                     Math.PI * (38 * 0.305) ^ 2 * (22 * 0.305) / 4) + ((1 / ((Solido / 100) / 2.7 + (1 - Solido / 100))) *                                                                  (60 / 100) * 0.4 * (jc / 100) * Math.PI * (38 * 0.305) ^ 2 * (22 * 0.305) / 4)) / ((jc / 100) * Math.PI * (38 * 0.305) ^ 2 * (22 * 0.305) / 4)) * (jc / 100 - 1.065 * jc * jc / 10000) * Math.Sin(39.9 * Math.PI / 180)) / (1 - 10 / 100))
            model1.AddConstraints("one", potencia == (0.238 * Math.Pow(diametro, 3.5) * (largo / (double)diametro) * (millSpeed / (double)100) * ((((1 - 0.4) * densidad_bolas * (y / (double)100) * Math.PI * Math.Pow((diametro * 0.305), 2) * (largo * 0.305) / (double)4) + ((1 / (double)((solidos / (double)100) / densidad_mineral + (1 - solidos / (double)100))) * (jci / (double)100 - y / (double)100) * Math.PI * Math.Pow((diametro * 0.305), 2) * (largo * 0.305) / (double)4) + ((1 / (double)((solidos / (double)100) / densidad_mineral + (1 - solidos / (double)100))) * (inter / (double)100) * 0.4 * (jci / (double)100) * Math.PI * Math.Pow((diametro * 0.305), 2) * (largo * 0.305) / (double)4)) / ((jci / (double)100) * Math.PI * Math.Pow((diametro * 0.305), 2) * (largo * 0.305) / (double)4)) * (jci / (double)100 - 1.065 * jci * jci / (double)10000) * Math.Sin(angulo * Math.PI / 180)) / (1 - losses / (double)100));

            SimplexDirective simplex = new SimplexDirective();

            Directive n = new Directive();
            n.TimeLimit = 300000;
            n.WaitLimit = 420000;
            //simplex.IterationLimit = 100;

            // Solve the problem
            Solution sol = context1.Solve(n);

            ///Report report = sol.GetReport();

            return y.GetDouble();

        }

    }
}
