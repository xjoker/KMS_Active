using System;
using System.Diagnostics;

namespace KMS_Active
{
    class Program
    {
        static void Main(string[] args)
        {
            string Server_address = ""; //激活服务器
            bool show_dlv = false; //是否显示激活情况

            if (args.Length < 2)
            {
                Console.WriteLine("Use examples\n");
                Console.WriteLine("\t Active Windows  \"KMS_Active.exe /w /s active.server.com \"");
                Console.WriteLine("\t Active Office  \"KMS_Active.exe /o /s active.server.com \"");
                Console.WriteLine("\t Active Windows and Office  \"KMS_Active.exe /w /o /s active.server.com \"");
                Console.WriteLine("");
                Console.WriteLine("\t  /s  Setting server address");
                Console.WriteLine("\t  /d  Display activation information");
                Console.WriteLine("\t  /o  Active office");
                Console.WriteLine("\t  /w  Active windows");
                Environment.Exit(0);
            }

            int server_address = Array.IndexOf<string>(args, "/s");
            int office_active = Array.IndexOf<string>(args, "/o");
            int windows_active = Array.IndexOf<string>(args, "/w");
            int dly_flag = Array.IndexOf<string>(args, "/d");

            if (server_address != -1)
            {
                Server_address = args[server_address + 1];
                if (Server_address == "")
                {
                    Console.WriteLine("Active server error!");
                    Environment.Exit(0);
                }

            }
            if (dly_flag != -1)
            {
                show_dlv = true;
            }

            // Windows active
            if (windows_active!=-1)
            {
                string SN = System_Info.KMS_key;
                if (SN != "")
                {
                    Process p = new Process();
                    p.StartInfo.FileName = @"cscript";

                    //导入序列号
                    Console.WriteLine("Import the licensing...");
                    p.StartInfo.Arguments = @"//B //Nologo " + Environment.SystemDirectory + @"\slmgr.vbs /ipk " + SN;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    p.Start();
                    p.WaitForExit();
                    p.Close();


                    Console.WriteLine("Set the Windows activation server...");
                    p.StartInfo.Arguments = @"//B //Nologo " + Environment.SystemDirectory + @"\slmgr.vbs /skms " + Server_address;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    p.Start();
                    p.WaitForExit();
                    p.Close();

                    Console.WriteLine("Activation Windows start...");
                    p.StartInfo.Arguments = @"//B //Nologo " + Environment.SystemDirectory + @"\slmgr.vbs /ato";
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    p.Start();
                    p.WaitForExit();
                    p.Close();
                    p.Dispose();

                }
                else
                {
                    Console.WriteLine("This Windows version no support!");
                }
            }

            if (office_active != -1)
            {
                string ospp = System_Info.Office_KMS_key;
                if (ospp != "")
                {
                    Process p = new Process();
                    p.StartInfo.FileName = @"cscript";

                    Console.WriteLine("Set the Office activation server...");
                    p.StartInfo.Arguments = @"//B //Nologo " + ospp + @" /sethst:" + Server_address;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    p.Start();
                    p.WaitForExit();
                    p.Close();


                    Console.WriteLine("Activation Office start...");
                    p.StartInfo.Arguments = @"//B //Nologo " + ospp + @" /act";
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    p.Start();
                    p.WaitForExit();
                    p.Close();
                }
                else
                {
                    Console.WriteLine("Not find Office !");
                }

            }

            if (show_dlv)
            {
                Process p = new Process();
                Console.WriteLine("Show dlv...");
                p.StartInfo.FileName = @"cmd.exe";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardInput = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.Arguments = "/C slmgr.vbs /dlv";
                p.Start();
                p.WaitForExit();
            }

        }
    }
}
