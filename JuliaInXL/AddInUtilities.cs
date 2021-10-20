using System;
using System.Collections;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;
using ExcelDna.ComInterop;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Registration;
using ExcelDna.Logging;

using NetMQ;
using Newtonsoft.Json.Linq;
using NetMQ.Sockets;
using Microsoft.Win32;
using System.Collections.Generic;

namespace JuliaInXL
{

    [Serializable()]
    public class JuliaException : System.Exception
    {
        public JuliaException() : base() { }
        public JuliaException(string message) : base(message) { }
        public JuliaException(string message, System.Exception inner) : base(message, inner) { }

        // A constructor is needed for serialization when an
        // exception propagates from a remoting server to the client.
        protected JuliaException(System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
        { }
    }

    [ComVisible(true)]
    public interface IAddInUtilities
    {
        object CallJulia(string fn, object[] arguments,int JuliaInXL_timeout=30);
        void Terminate();
        void Reconnect();
        Process LaunchLocalJulia();
        void ShutdownJulia(Process p);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class AddInUtilities : IAddInUtilities
    {
        private RequestSocket client;
        private NetMQContext context;
        private Process juliaprocess = new Process();
        private string juliaversion;
        private string juliaproversion;
        private string endpoint = "tcp://localhost:9999";
        private string juliafile = "";
        private bool connected = false;

        public string Endpoint
        {
            get
            {
                return endpoint;
            }

            set
            {
                endpoint = value;
            }
        }

        public Process JuliaProcess
        {
            get
            {
                return juliaprocess;
            }

            set
            {
                juliaprocess = value;
            }
        }

        public string JuliaProVersion
        {
            get
            {
                return juliaproversion;
            }

            set
            {
                juliaproversion = value;
            }
        }

        public string JuliaVersion
        {
            get
            {
                return juliaversion;
            }

            set
            {
                juliaversion = value;
            }
        }

        public string JuliaFile
        {
            get
            {
                return juliafile;
            }
            set
            {
                juliafile = value;
            }
        }

        public bool Connected
        {
            get
            {
                return connected;
            }
            set
            {
                connected = value;
            }
        }

        public static bool registryValueExists(string hive_HKLM_or_HKCU, string registryRoot, string valueName)
        {
            RegistryKey root;

            switch (hive_HKLM_or_HKCU.ToUpper())
            {
                case "HKLM":
                    RegistryKey localMachineRegistry
                     = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                                      Environment.Is64BitOperatingSystem
                                          ? RegistryView.Registry64
                                          : RegistryView.Registry32);

                    root = localMachineRegistry.OpenSubKey(registryRoot);

                    break;
                case "HKCU":

                    root = Registry.CurrentUser.OpenSubKey(registryRoot);

                    break;
                default:
                    throw new System.InvalidOperationException("parameter registryRoot must be either \"HKLM\" or \"HKCU\"");
            }

            bool test = (root != null) && (root.GetValue(valueName) != null);

            return test;
        }

        public AddInUtilities()
        {
            
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            JuliaVersion = version.Major + "." + version.Minor + "." + version.Build;
            JuliaProVersion = JuliaVersion + "." + version.Revision;

            // Check for either JuliaInXL or JuliaPro contained within the registry path to the JuliaPro installation.
            if (registryValueExists("HKCU", "Software\\JuliaPro\\JuliaInXL", "JuliaInXL_Default_Endpoint") == true)
            {
                endpoint = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Software\JuliaPro\JuliaInXL", "JuliaInXL_Default_Endpoint", "");
            }
            else if (registryValueExists("HKCU", "Software\\JuliaPro\\" + JuliaProVersion, "JuliaInXL_Default_Endpoint") == true)
            {
                endpoint = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Software\JuliaPro\" + JuliaProVersion, "JuliaInXL_Default_Endpoint", "");
            }
            else if (registryValueExists("HKLM", "Software\\JuliaPro\\JuliaInXL", "JuliaInXL_Default_Endpoint") == true)
            {
                RegistryKey localMachineRegistry
                     = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                                      Environment.Is64BitOperatingSystem
                                          ? RegistryView.Registry64
                                          : RegistryView.Registry32);

                RegistryKey key = localMachineRegistry.OpenSubKey("Software\\JuliaPro\\JuliaInXL");

                endpoint = (string)key.GetValue("JuliaInXL_Default_Endpoint");
            }
            else if (registryValueExists("HKLM", "Software\\JuliaPro\\" + JuliaProVersion, "JuliaInXL_Default_Endpoint") == true)
            {
                RegistryKey localMachineRegistry
                     = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                                      Environment.Is64BitOperatingSystem
                                          ? RegistryView.Registry64
                                          : RegistryView.Registry32);

                RegistryKey key = localMachineRegistry.OpenSubKey("Software\\JuliaPro\\" + JuliaProVersion);

                endpoint = (string)key.GetValue("JuliaInXL_Default_Endpoint");
            }
            else if (Environment.GetEnvironmentVariable("JULIAINXL_DEFAULT_ENDPOINT") != null)
            {
                endpoint = Environment.GetEnvironmentVariable("JULIAINXL_DEFAULT_ENDPOINT");
            }
            else
            {
                endpoint = "tcp://localhost:9999";
            }

            context = NetMQ.NetMQContext.Create();
            client = context.CreateRequestSocket();
            client.Connect(endpoint);
            Connected = true;
        }

        public Process LaunchLocalJulia()
        {
            string julia_exe;
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            JuliaVersion = version.Major + "." + version.Minor + "." + version.Build;
            JuliaProVersion = JuliaVersion + "." + version.Revision;

            // Check for either JuliaInXL or JuliaPro contained within the registry path to the JuliaPro installation.
            if (registryValueExists("HKCU", "Software\\JuliaPro\\JuliaInXL\\", "Install_Dir") == true)
            {
                string juliaBaseDir = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Software\JuliaPro\JuliaInXL\", "Install_Dir", "");
                julia_exe = juliaBaseDir + "\\julia.exe";
            }
            else if (registryValueExists("HKCU", "Software\\JuliaPro\\" + JuliaProVersion, "Install_Dir") == true)
            {
                string juliaBaseDir = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Software\JuliaPro\" + JuliaProVersion, "Install_Dir", "");
                julia_exe = juliaBaseDir + "\\Julia-" + JuliaVersion + "\\bin\\julia.exe";
            }
            else if (registryValueExists("HKLM", "Software\\JuliaInXL\\" + JuliaProVersion, "Install_Dir") == true)
            {
                RegistryKey localMachineRegistry
                     = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                                      Environment.Is64BitOperatingSystem
                                          ? RegistryView.Registry64
                                          : RegistryView.Registry32);
                RegistryKey key = localMachineRegistry.OpenSubKey("Software\\JuliaPro\\JuliaInXL");
                string juliaBaseDir = (string) key.GetValue("Install_Dir");
                julia_exe = juliaBaseDir + "\\julia.exe";
            }
            else if (registryValueExists("HKLM", "Software\\JuliaPro\\" + JuliaProVersion, "Install_Dir") == true)
            {
                RegistryKey localMachineRegistry
                     = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                                      Environment.Is64BitOperatingSystem
                                          ? RegistryView.Registry64
                                          : RegistryView.Registry32);
                RegistryKey key = localMachineRegistry.OpenSubKey("Software\\JuliaPro\\" + JuliaProVersion);
                string juliaBaseDir = (string)key.GetValue("Install_Dir");
                julia_exe = juliaBaseDir + "\\Julia-" + JuliaVersion + "\\bin\\julia.exe";
            }
            else
            {
                julia_exe = "julia.exe";
            }

            Process p = new Process();

            try
            {
                string endpoint = Endpoint;
                char[] delimiter = { ':' };
                string[] words = endpoint.Split(delimiter);
                string port = words[words.Length - 1];
                string host = words[words.Length - 2];

                string address;
                if (host.Equals("//localhost"))
                {
                    address = "//127.0.0.1";
                }
                else
                {
                    address = host;
                }

                string commandlineargs = String.Join("", "-i -e \"using JuliaInXL; JuliaInXL.start_async_server(", port, ")\"");

                p.StartInfo.FileName = julia_exe;
                p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized; // ProcessWindowStyle.Hidden;
                p.StartInfo.Arguments = commandlineargs;

                p.Start();

                Thread.Sleep(1000);
                Process.GetProcessById(p.Id);
                Connected = true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception = " + e.ToString());
                p = null;
            }

            return p;

        }

        public void ShutdownJulia(Process p)
        {
            p.Kill();
        }


        public object CallJulia(string fn, object[] arguments, int JuliaInXL_timeout=30)
        {

            try
            {
                Version version = Assembly.GetExecutingAssembly().GetName().Version;
                JuliaVersion = version.Major + "." + version.Minor + "." + version.Build;
                JuliaProVersion = JuliaVersion + "." + version.Revision;
                RegistryKey localMachineRegistry
                         = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                                          Environment.Is64BitOperatingSystem
                                              ? RegistryView.Registry64
                                              : RegistryView.Registry32);
                RegistryKey key = localMachineRegistry.OpenSubKey("Software\\JuliaPro\\" + JuliaProVersion + "\\JuliaInXL");

                if ((registryValueExists("HKCU", "Software\\JuliaPro\\" + JuliaProVersion + "\\JuliaInXL", "Timeout") == true) && (int.TryParse((string)Registry.GetValue(@"HKEY_CURRENT_USER\Software\JuliaPro\" + JuliaProVersion + "\\JuliaInXL", "Timeout", ""), out JuliaInXL_timeout)) && (JuliaInXL_timeout > 0)) {}               
                else if ((registryValueExists("HKLM", "Software\\JuliaPro\\" + JuliaProVersion + "\\JuliaInXL", "Timeout") == true) && (int.TryParse((string)key.GetValue("Timeout"), out JuliaInXL_timeout)) && (JuliaInXL_timeout > 0)) {}
                else if (Environment.GetEnvironmentVariable("JuliaInXL_timeout") != null && (int.TryParse(Environment.GetEnvironmentVariable("JuliaInXL_timeout"), out JuliaInXL_timeout)) && (JuliaInXL_timeout > 0)) {}
                else
                {
                    JuliaInXL_timeout = 30;
                }
                JObject o = new JObject();
                o["cmd"] = fn;
                if (arguments.Length > 1 || arguments[0] != null)
                {
                    JArray args = new JArray();
                    for (int i = 0; i < arguments.Length; i++)
                    {

                        if (arguments[i] is object[])
                        {

                            object[] arg = (object[])arguments[i];

                            JArray jarg = new JArray();
                            for (int k = 0; k < arg.Length; k++)
                            {
                                jarg.Add(arg[k]);
                            }

                            args.Add(jarg);

                        }
                        else if (arguments[i] is object[,])
                        {
                            object[,] arg = (object[,])arguments[i];

                            int cols = arg.GetUpperBound(1) + 1;
                            int rows = arg.GetUpperBound(0) + 1;

                            JArray jarg_cols = new JArray();
                            for (int n = 0; n < cols; n++)
                            {
                                object[] arg_col = new object[rows];
                                for (int m = 0; m < rows; m++)
                                {
                                    arg_col[m] = arg[m, n];
                                }
                                jarg_cols.Add(new JArray(arg_col));
                            }

                            args.Add(jarg_cols);

                        }
                        else if (arguments[i] is Array)
                        {
                            //It seems JArray flattens its arguments by default
                            //wrapping it in another JArray prevents flattening
                            args.Add(new JArray(arguments[i]));
                        }
                        else
                        {
                            args.Add(arguments[i]);
                        }
                    }
                    o["args"] = args;
                }                
                client.SendFrame(o.ToString());

                // Receive the response from the client socket
                string m2="";
                if (Connected==true)
                { 
                client.TryReceiveFrameString(TimeSpan.FromSeconds(JuliaInXL_timeout), out m2);                
                }
                object retval = ProcessResult(m2);                
                return retval;
            }
            catch
            {
                this.Reconnect();
                return "#ERR";
            }
        }

        /**
        * Convert the JSON string into native objects that can be returned to Excel
        * Arrays are supported only upto dimension 2
        **/
        private dynamic ProcessResult(string m2)
        {
            dynamic retval = "";
            if (m2 == null)
            {
                client.Close();
                client = context.CreateRequestSocket();
                client.Connect(endpoint);
                Connected = false;
                retval = "#ERR";
            }
            else
            {
                JObject result = JObject.Parse(m2);
                retval = result["data"];

                if (retval is JArray)
                {
                    Object[] t = ((JArray)retval).ToObject<Object[]>();
                    if ((t).Length > 0 && t[0] is JArray)
                    {
                        int cols = (t).Length;
                        Object[] col = ((JArray)t[0]).ToObject<Object[]>();
                        int rows = (col).Length;
                        object[,] r = new object[rows, cols];
                        //Array is two Dimensional
                        for (int j = 0; j < cols; j++)
                        {
                            col = ((JArray)t[j]).ToObject<Object[]>();
                            for (int i = 0; i < rows; i++)
                            {
                                r[i, j] = col[i];
                            }
                            t[j] = ((JArray)t[j]).ToObject<object[]>();
                        }
                        retval = r;

                    }
                    else //if ((t).Length > 0 && t[0] is double)
                    {
                        object[] r = new object[(t).Length];
                        for (int i = 0; i < t.Length; i++)
                        {
                            r[i] = t[i];
                        }
                        retval = r;
                    }
                }

            }

            return retval;
        }

        public void Terminate()
        {
            JObject o = new JObject();
            o["cmd"] = ":terminate";
            client.SendFrame(o.ToString());
            this.Reconnect();
        }

        public void Reconnect()
        {
            client.Close();
            client = context.CreateRequestSocket();
            client.Connect(endpoint);
            Connected = true;

        }

        public void KeepAlive()
        {
            string KeepAlive_endpoint = Endpoint;
            char[] KeepAlive_delimiter = { ':' };
            string[] KeepAlive_words = KeepAlive_endpoint.Split(KeepAlive_delimiter);
            string KeepAlive_port = KeepAlive_words[KeepAlive_words.Length - 1];
            string KeepAlive_host = KeepAlive_words[KeepAlive_words.Length - 2];
            int KeepAlive_Interger_port = (Int32.Parse(KeepAlive_port)) - 1;
            string KeepAlive_connection_string = $"tcp:{KeepAlive_host}:{KeepAlive_Interger_port}";
			try
			{
				while (true)
				{
						using (var keep_alive_client = new RequestSocket())  
						{	
							keep_alive_client.Connect(KeepAlive_connection_string);							
							keep_alive_client.SendFrame("Hello");							
							string keep_alive_sampletext;
							keep_alive_client.TryReceiveFrameString(TimeSpan.FromSeconds(1), out keep_alive_sampletext);							
							if (keep_alive_sampletext == "Hello")
							{
								Connected = true;
							}
							else
							{                        
								JObject o = new JObject();
								o["cmd"] = "JuliaInXL.start_heartbeat";
								o["args"] = KeepAlive_Interger_port;
								client.SendFrame(o.ToString());
								string m2;
								client.TryReceiveFrameString(TimeSpan.FromSeconds(1), out m2);
								object retval = ProcessResult(m2);                       
								if (retval.ToString() == "#ERR" || retval.ToString() == "")
								{
									Connected = false;
								}
								else
								{
									Connected = true;
								}
							}
							// Making the thread wait
							Thread.CurrentThread.Join(1000);							
						}
				}            
			}
			catch {}
        }

        public void close()
        {
            if (client != null) { client.Close(); }
            if (context != null) { try { context.Terminate(); } catch { } }
        }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("JuliaInXLutilities")]
    public class JuliaInXLutilities
    {

        private AddInUtilities utilities;

        public AddInUtilities Utilities
        {
            get
            {
                return utilities;
            }

            set
            {
                utilities = value;
            }

        }

        public object LaunchUtilities()
        {
            if (utilities == null)
                utilities = new AddInUtilities();

            return utilities;
        }

    }

    public static class JuliaInXLutilities_global
    {
        static JuliaInXLutilities _jxlu;

        public static JuliaInXLutilities jxlu
        {
            get
            {
                return _jxlu;
            }

            set
            {
                _jxlu = value;
            }
        }
    }

    [ComVisible(true)]
    [ProgId("JuliaComAddIn")]
    public class JuliaComAddIn : ExcelComAddIn
    {

        public JuliaComAddIn()
        {
        }
        public override void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
        }
        public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
        }
        public override void OnAddInsUpdate(ref Array custom)
        {
        }
        public override void OnStartupComplete(ref Array custom)
        {
            JuliaInXLutilities_global.jxlu.Utilities.JuliaProcess = JuliaInXLutilities_global.jxlu.Utilities.LaunchLocalJulia();
            Thread thread = new Thread(new ThreadStart(JuliaInXLutilities_global.jxlu.Utilities.KeepAlive));
            thread.Start();                     
        }
        public override void OnBeginShutdown(ref Array custom)
        {
            JuliaInXLutilities_global.jxlu.Utilities.ShutdownJulia(JuliaInXLutilities_global.jxlu.Utilities.JuliaProcess);
        }
    }

    [ComVisible(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("JuliaAddIn")]
    public class JuliaAddIn : IExcelAddIn
    {

        ExcelDna.Integration.ExcelComAddIn _comAddIn;

        public void AutoOpen()
        {
            JuliaInXLutilities_global.jxlu = new JuliaInXLutilities();
            JuliaInXLutilities_global.jxlu.LaunchUtilities();
            ExcelRegistration.GetExcelFunctions().
               ProcessParamsRegistrations().
               RegisterFunctions();
            ComServer.DllRegisterServer();
            _comAddIn = new JuliaComAddIn();
            ExcelComAddInHelper.LoadComAddIn(_comAddIn);
            ExcelIntegration.RegisterUnhandledExceptionHandler(
                delegate (object ex) {
                    Exception ex2 = (Exception)ex;
                    return string.Format("#{0}!", ex2.Message);
                }
            );
        }

        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }

        [ExcelFunction(ExplicitRegistration = true, IsMacroType = false)]
        public static object[,] jlcall(
                                    [ExcelArgument(Name = "function", Description = "Name of Julia function", AllowReference = true)] object fn,
                                    [ExcelArgument(Name = "arguments", Description = "List input arguments to Julia function", AllowReference = true)] params object[] args)

        {
            bool xlErrorPresent = false;
            ExcelReference caller = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            int rows = caller.RowLast - caller.RowFirst + 1;
            int cols = caller.ColumnLast - caller.ColumnFirst + 1;
            ExcelReference cellRef, a;
            dynamic retval;
            object[,] ret;
            object b;
            ExcelError err = ExcelError.ExcelErrorValue;

            // List of numeric types used later for comparison
            List<Type> numerictypes = new List<Type>();
            numerictypes.Add(typeof(double));
            numerictypes.Add(typeof(float));
            numerictypes.Add(typeof(int));
            numerictypes.Add(typeof(Int16));
            numerictypes.Add(typeof(Int32));
            numerictypes.Add(typeof(Int64));
            numerictypes.Add(typeof(UInt16));
            numerictypes.Add(typeof(UInt32));
            numerictypes.Add(typeof(UInt64));

            string f;

            if (JuliaInXLutilities_global.jxlu.Utilities.Connected)
            {
                try
                {
                    if (fn is ExcelReference)
                    {
                        f = ((ExcelReference)fn).GetValue() as string;
                    }
                    else
                    {
                        f = (string)fn;
                    }

                    int numargs = args.Length;
                    if (numargs > 0)
                    {
                        object[] jargs = new object[numargs];
                        for (int i = 0; i < numargs; i++)
                        {
                            if (args[i] is ExcelReference)
                            {
                                a = (ExcelReference)args[i];
                                int rf = a.RowFirst;
                                int rl = a.RowLast;
                                int cf = a.ColumnFirst;
                                int cl = a.ColumnLast;
                                int rowsarg = rl - rf + 1;
                                int colsarg = cl - cf + 1;
                                if (rl == rf && cl == cf)
                                {
                                    cellRef = new ExcelReference(rf, rf, cf, cf, a.SheetId);
                                    b = cellRef.GetValue();
                                    if (b is ExcelEmpty)
                                    {
                                        throw new JuliaException("JuliaEmptyCell");
                                    }
                                    if (b is ExcelError)
                                    {
                                        xlErrorPresent = true;
                                        err = (ExcelError)b;
                                        break;
                                    }
                                    jargs[i] = b;
                                }
                                else if (rl == rf && cl != cf)
                                {
                                    object[] arg = new object[colsarg];
                                    for (int n = 0; n < colsarg; n++)
                                    {
                                        cellRef = new ExcelReference(rf, rf, cf + n, cf + n, a.SheetId);
                                        b = cellRef.GetValue();
                                        if (b is ExcelEmpty)
                                        {
                                            throw new JuliaException("JuliaEmptyCell");
                                        }
                                        if (b is ExcelError)
                                        {
                                            xlErrorPresent = true;
                                            err = (ExcelError) b;
                                            break;
                                        }
                                        arg[n] = b;
                                    }
                                    jargs[i] = arg;
                                }
                                else if (rl != rf && cl == cf)
                                {
                                    object[] arg = new object[rowsarg];
                                    for (int m = 0; m < rowsarg; m++)
                                    {
                                        cellRef = new ExcelReference(rf + m, rf + m, cf, cf, a.SheetId);
                                        b = cellRef.GetValue();
                                        if (b is ExcelDna.Integration.ExcelEmpty)
                                        {
                                            throw new JuliaException("JuliaEmptyCell");
                                        }
                                        if (b is ExcelError)
                                        {
                                            xlErrorPresent = true;
                                            err = (ExcelError) b;
                                            break;
                                        }

                                        arg[m] = b;
                                    }
                                    if (xlErrorPresent == false)
                                    {
                                        jargs[i] = arg;
                                    }
                                }
                                else
                                {
                                    object[,] arg = new object[rowsarg, colsarg];
                                    for (int m = 0; m < rowsarg; m++)
                                    {
                                        for (int n = 0; n < colsarg; n++)
                                        {
                                            cellRef = new ExcelReference(rf + m, rf + m, cf + n, cf + n, a.SheetId);
                                            b = cellRef.GetValue();
                                            if (b is ExcelEmpty)
                                            {
                                                throw new JuliaException("JuliaEmptyCell");
                                            }
                                            if (b is ExcelError)
                                            {
                                                xlErrorPresent = true;
                                                err = (ExcelError) b;
                                                break;
                                            }
                                            arg[m, n] = b;
                                        }
                                    }
                                    if (xlErrorPresent == false)
                                    {
                                        jargs[i] = arg;
                                    }
                                }
                            }
                            else
                            {

                                if (args[i] is ExcelEmpty)
                                {
                                    throw new JuliaException("JuliaEmptyCell");
                                }
                                if (args[i] is ExcelError)
                                {
                                    xlErrorPresent = true;
                                    err = (ExcelError)args[i];
                                    break;
                                }

                                jargs[i] = args[i];
                            }
                        }

                        if (xlErrorPresent == false)
                        {
                            retval = JuliaInXLutilities_global.jxlu.Utilities.CallJulia(f, jargs);
                        }
                        else
                        {
                            retval = new object[rows, cols];
                            for (int i = 0; i < rows; ++i)
                            {
                                for (int j = 0; j < cols; ++j)
                                {
                                    retval[i, j] = err;
                                }
                            }
                        }
                    }
                    else
                    {
                        object[] zargs = new object[1];
                        zargs[0] = null;
                        retval = JuliaInXLutilities_global.jxlu.Utilities.CallJulia(f, zargs);
                    }

                    if (retval is object[])
                    {
                        // Size of the data returned by Julia
                        int M = ((object[])retval).GetUpperBound(0) + 1;

                        // rows and cols are the size of the active cells in Excel
                        ret = new object[rows, cols];
                        object r;
                        for (int i = 0; i < rows; ++i)
                        {
                            for (int j = 0; j < cols; ++j)
                            {
                                int k = j * rows + i;
                                r = retval[k];
                                if (k < M)
                                {
                                    if (r is JValue)
                                    {
                                        Type t = ((JValue) r).Value.GetType();
                                        if (numerictypes.Contains(t))
                                        {
                                            ret[i, j] = (double) r;
                                        }
                                        else if (t == typeof(bool))
                                        {
                                            ret[i, j] = (bool) r;
                                        }
                                        else if (t == typeof(string))
                                        {
                                            ret[i, j] = (String) r;
                                        }
                                        else
                                        {
                                            ret[i, j] = r;
                                        }
                                    }
                                    else
                                    {
                                        ret[i, j] = r;
                                    }
                                }
                                else
                                {
                                    ret[i, j] = ExcelError.ExcelErrorNA;
                                }
                            }
                        }
                        return ret;
                    }
                    else if (retval is object[,])
                    {
                        // Size of the data returned by Julia
                        int M = ((object[,])retval).GetUpperBound(0) + 1;
                        int N = ((object[,])retval).GetUpperBound(1) + 1;

                        ret = new object[rows, cols];
                        object r;
                        for (int i = 0; i < rows; ++i)
                        {
                            for (int j = 0; j < cols; ++j)
                            {
                                if (i < M && j < N)
                                {
                                    r = retval[i, j];
                                    if (r is JValue)
                                    {
                                        Type t = ((JValue)r).Value.GetType();
                                        if (numerictypes.Contains(t))
                                        {
                                            ret[i, j] = (double) r;
                                        }
                                        else if (t == typeof(bool))
                                        {
                                            ret[i, j] = (bool) r;
                                        }
                                        else if (t == typeof(string))
                                        {
                                            ret[i, j] = (String) r;
                                        }
                                        else
                                        {
                                            ret[i, j] = r;
                                        }
                                    }
                                    else
                                    {
                                        ret[i, j] = r;
                                    }
                                }
                                else
                                {
                                    ret[i, j] = ExcelError.ExcelErrorNA;
                                }
                            }
                        }

                        return ret;
                    }
                    else if (retval is JValue)
                    {
                        Type t = retval.Value.GetType();
                        if (numerictypes.Contains(t))
                        {
                            ret = new object[rows, cols];
                            ret[0, 0] = (double) retval;
                            return ret;
                        }
                        else if (t == typeof(bool))
                        {
                            ret = new object[rows, cols];
                            ret[0, 0] = (bool) retval;
                            return ret;
                        }
                        else if (t == typeof(string))
                        {
                            ret = new object[rows, cols];
                            ret[0, 0] = (String) retval;
                            return ret;
                        }
                    }
                    return retval;
                }
                catch (JuliaException e)
                {
                    throw e;
                }
                catch
                {
                    retval = new object[rows, cols];
                    for (int i = 0; i < rows; ++i)
                    {
                        for (int j = 0; j < cols; ++j)
                        {
                            retval[i, j] = ExcelError.ExcelErrorValue;
                        }
                    }
                    return retval;
                }
            }
            else
            {
                throw new JuliaException("JuliaNotConnected");
            }
        }

        [ExcelFunction(ExplicitRegistration = true, IsMacroType = false)]
        public static object[,] jleval([ExcelArgument(Name = "argument", Description = "String of Julia expression to be executed by eval(parse(expr))", AllowReference = true)] string arg)
        {
            object[] args = { arg };
            return jlcall("parse_and_eval", args);
        }
        
        [ExcelFunction(ExplicitRegistration = true, IsMacroType = false)]
        public static object[,] jlsetvar([ExcelArgument(Name = "variable", Description = "Name of a global variable to set in the Julia process", AllowReference = true)] object v,
                                         [ExcelArgument(Name = "arguments", Description = "One or more values to assign to this global variable", AllowReference = true)] params object[] args)
        {
            object[] a = new object[args.Length + 1];
            a[0] = v;
            for (int i = 0; i< args.Length; ++i)
            {
                a[i + 1] = args[i];
            }
            return jlcall("jlsetvar", a);
        }
        
    }

    

    [ComVisible(true)]
    public class JuliaRibbon : ExcelRibbon
    {

        private IRibbonUI ribbon = null;

        private IRibbonUI Ribbon
        {
            get
            {
                return ribbon;
            }

            set
            {
                ribbon = value;
            }
        }

        public void OnLoad_JuliaInXL(IRibbonUI ribbon)
        {
            this.ribbon = ribbon;
            JuliaInXLutilities_global.jxlu.LaunchUtilities();
            ribbon.Invalidate();
        }

        public void OnButtonLaunchJulia_JuliaInXL(IRibbonControl control)
        {
            char[] delimiter = { ':' };
            string[] words = JuliaInXLutilities_global.jxlu.Utilities.Endpoint.Split(delimiter);
            if (!words[1].Equals("//localhost"))
            {
                words[1] = "//localhost";
                JuliaInXLutilities_global.jxlu.Utilities.Endpoint = words[0] + ":" + words[1] + ":" + words[2];
                ribbon.Invalidate();
            }

            if (words[0].Equals("tcp") && words[1].Equals("//localhost"))
            {
                if (JuliaInXLutilities_global.jxlu.Utilities.JuliaProcess != null)
                {
                    try
                    {
                        Process.GetProcessById(JuliaInXLutilities_global.jxlu.Utilities.JuliaProcess.Id);
                        JuliaInXLutilities_global.jxlu.Utilities.JuliaProcess.Kill();
                        JuliaInXLutilities_global.jxlu.Utilities.JuliaProcess = null;
                    }
                    catch (Exception)
                    {
                        JuliaInXLutilities_global.jxlu.Utilities.JuliaProcess = null;
                    }
                }
                JuliaInXLutilities_global.jxlu.Utilities.JuliaProcess = JuliaInXLutilities_global.jxlu.Utilities.LaunchLocalJulia();
            }

        }

        public String GetEndpointText_JuliaInXL(IRibbonControl control)
        {
            return JuliaInXLutilities_global.jxlu.Utilities.Endpoint;
        }

        public String GetJuliaFileText_JuliaInXL(IRibbonControl control)
        {
            return JuliaInXLutilities_global.jxlu.Utilities.JuliaFile;
        }

        public void SetConnectionInfo_JuliaInXL(IRibbonControl control, string text)
        {
            //TODO: Determine proper string validation for an endpoint.

            char[] delimiter = { ':' };
            string[] words = text.Split(delimiter);
            string portdigits = words[2];
            Regex rgx = new Regex("^\\d{" + portdigits.Length.ToString() + "}");
            if (words.Length == 3 &&
                words[0].Equals("tcp") &&
                words[1].Substring(0, 2).Equals("//") &&
                rgx.IsMatch(portdigits))
            {
                JuliaInXLutilities_global.jxlu.Utilities.Endpoint = text;
                JuliaInXLutilities_global.jxlu.Utilities.Reconnect();
            }

            MessageBox.Show("Current endpoint " + JuliaInXLutilities_global.jxlu.Utilities.Endpoint);
        }

        public void SetJuliaFile_JuliaInXL(IRibbonControl control, string text)
        {
            JuliaInXLutilities_global.jxlu.Utilities.JuliaFile = text;
        }

        public void OnButtonReconnect_JuliaInXL(IRibbonControl control)
        {
            JuliaInXLutilities_global.jxlu.Utilities.Reconnect();
        }

        public static void OnButtonTerminate_JuliaInXL()
        {
            JuliaInXLutilities_global.jxlu.Utilities.Terminate();
        }

        public void OnButtonSelectJuliaFile_JuliaInXL(IRibbonControl control)
        {
            string juliaVersion = JuliaInXLutilities_global.jxlu.Utilities.JuliaVersion;
            string juliaProVersion = JuliaInXLutilities_global.jxlu.Utilities.JuliaProVersion;
            string initialDirectory;

            if (AddInUtilities.registryValueExists("HKCU", "Software\\JuliaPro\\JuliaInXL", "Install_Dir") == true)
            {
                string juliaBaseDir = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Software\JuliaPro\JuliaInXL", "Install_Dir", "");
                initialDirectory = juliaBaseDir;
            }
            else if (AddInUtilities.registryValueExists("HKLM", "Software\\JuliaPro\\JuliaInXL", "Install_Dir") == true)
            {
                RegistryKey localMachineRegistry
                     = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                                      Environment.Is64BitOperatingSystem
                                          ? RegistryView.Registry64
                                          : RegistryView.Registry32);
                RegistryKey key = localMachineRegistry.OpenSubKey("Software\\JuliaPro\\JuliaInXL");
                string juliaBaseDir = (string)key.GetValue("Install_Dir");

                initialDirectory = juliaBaseDir;
            }
            else
            {
                initialDirectory = Environment.GetEnvironmentVariable("HOMEDRIVE") + "\\" + Environment.GetEnvironmentVariable("HOMEPATH");
            }

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = initialDirectory;
            openFileDialog1.Filter = "Julia files (*.jl)|*.jl|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.RestoreDirectory = true;
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                JuliaInXLutilities_global.jxlu.Utilities.JuliaFile = openFileDialog1.FileName;
                this.ribbon.InvalidateControl("JuliaSelectedFile");
            }
            
        }

        public void OnButtonIncludeJuliaFile_JuliaInXL(IRibbonControl control)
        {
            string file = (string)JuliaInXLutilities_global.jxlu.Utilities.JuliaFile.Clone();
            object[] arg = { file };
            if (System.IO.File.Exists(file) && JuliaInXLutilities_global.jxlu.Utilities.Connected == true)
            {
                    JuliaInXLutilities_global.jxlu.Utilities.CallJulia("include", arg);
            }
        }

        public override string GetCustomUI(string RibbonID)
        {

            string customUIXml =
                @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' loadImage='LoadImage' onLoad='OnLoad_JuliaInXL'>
                    <ribbon>
                    <tabs>
                        <tab id='CustomTab' label='JULIA'>
                        <group id='SampleGroup' label='Julia'>
                            <button id='JuliaButton' label='Launch Local Julia' image='juliaicon' size='large' onAction='OnButtonLaunchJulia_JuliaInXL' supertip='Launches a julia.exe process with JuliaInXL loaded by default at the currently configured endpoint.  The julia.exe process must be associated with a JuliaPro installation, and is determined either from the Windows registry or from the Path environment variable.'/>
                            <editBox id='JuliaSelectedFile' sizeString='xx                 xx               xx           xx'  image='juliafileselection' onChange='SetJuliaFile_JuliaInXL' maxLength='1024' getText='GetJuliaFileText_JuliaInXL' screentip='Julia File Path' supertip='Full path to the currently selected Julia source file to include in the associated Julia process.'/>                            
                            <button id='JuliaFileOpen' label='Select Julia File' image='juliafileopen' size='normal' onAction='OnButtonSelectJuliaFile_JuliaInXL' supertip='Open a file chooser dialog box to select a Julia source file to include in the associated Julia process.'/>
                            <button id='JuliaFileInclude' label='Include Julia File' image='juliafileinclude' size='normal' onAction='OnButtonIncludeJuliaFile_JuliaInXL' supertip='Load the currently selected Julia source file into the associated Julia process via the include command.'/>
                            <editBox id='ConnectionInfo' sizeString='tcp://localhost:9999999' imageMso='TracePrecedentsRemoveArrows' onChange='SetConnectionInfo_JuliaInXL' maxLength='1024' getText='GetEndpointText_JuliaInXL' screentip='JuliaInXL TCP endpoint' supertip='The TCP endpoint to use when connecting to the associated julia.exe process in which the JuliaInXL server is executing.'/>
                            <button id='ReconnectButton' label='Reconnect' imageMso='RefreshAll' size='normal' onAction='OnButtonReconnect_JuliaInXL' screentip='Reconnect to an existing Julia server session.' supertip='If a jlcall, jlsetvar, or jleval operation has previously failed to connect to a running julia server process, then the Reconnect button will reset the connection to that existing julia process.'/>
                            <button id='TerminateButton' label='Terminate' imageMso='RefreshCancel' size='normal' onAction='OnButtonTerminate_JuliaInXL'/>
                        </group >
                        </tab>
                    </tabs>
                    </ribbon>
                </customUI>";

            return customUIXml;
        }
    }
}
