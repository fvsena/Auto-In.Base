using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TP_Automation.Logger.Result;

namespace TPAuto_In.Base.Configuration
{
    public class Global
    {
        //FIXOS
        public static int AutomationId = 0;
        public static int ClientId = 0;
        public static int IdProcessoMarcacaoPonto = 0;
        public static int IdProcessoEncerramentoManual = 0;
        public static int IdProcessoXXXXX = 0;

        //APPCONFIG
        public static bool SkipCL = ConfigurationManager.AppSettings["SKIP_CL"].Equals("1");
        public static bool IsProd = ConfigurationManager.AppSettings["IS_PROD"].Equals("1");
        public static bool IsDebug = ConfigurationManager.AppSettings["IS_DEBUG"].Equals("1");
        public static string OverrideUser = ConfigurationManager.AppSettings["OVERRIDE_USER"];
        public static bool ExibeConsole = ConfigurationManager.AppSettings["CONSOLE"].Equals("1");
        public static bool TPLoginEnabled = ConfigurationManager.AppSettings["TP_LOGIN_ENABLED"].ToString().Equals("1");

        //CARREGADOS VIA AUTOMATION PARAMETER
        public static string Homologadores = null;
        public static string Celulas = null;
        public static string TipoLogin = null;
        public static string CaminhoIdentificadorTPLogin = null;
        public static int TempoEncerramento;
        public static bool LiberaEncerramentoManual;
        public static int TempoMaximoEsperaPonto;

        public static void ProcessarParametros(AutomationResult automation)
        {
            Homologadores = automation.Automation.Parameters.FirstOrDefault(x => x.Key.Equals("Homologadores")).Value;
            TempoEncerramento = Convert.ToInt32(automation.Automation.Parameters.FirstOrDefault(x => x.Key.Equals("Tempo Encerramento Manual")).Value);
            LiberaEncerramentoManual = automation.Automation.AutomationProcesses.FirstOrDefault(x => x.Name.Equals("Encerramento Manual")).Active;
            Celulas = automation.Automation.Parameters.FirstOrDefault(x => x.Key.Equals("Celulas com Auto-IN")).Value;
            TipoLogin = automation.Automation.Parameters.FirstOrDefault(x => x.Key.Equals("TP Login Habilitado")).Value;
            TempoMaximoEsperaPonto = Convert.ToInt32(automation.Automation.Parameters.FirstOrDefault(x => x.Key.Equals("Tempo Máximo Espera Controle Legal")).Value);
            CaminhoIdentificadorTPLogin = automation.Automation.Parameters.FirstOrDefault(x => x.Key.Equals("Identificador TP Login")).Value;
        }
    }
}
