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

            if (args.Length<2)
            {
                Console.WriteLine("Silent activation");
                Console.WriteLine("KMS_Active.exe /s active.server.com ");
                Console.WriteLine("");
                Console.WriteLine("Activation information is displayed");
                Console.WriteLine("KMS_Active.exe /d /s active.server.com");
                Environment.Exit(0);
            }

            int serviceTag = Array.IndexOf<string>(args, "/s");
            int dly_flag = Array.IndexOf<string>(args, "/d");
            if (serviceTag != -1)
            {
                Server_address = args[serviceTag + 1];
                if (Server_address == "")
                {
                    Console.WriteLine("Active server error!");
                    Environment.Exit(0);
                }

            }
            if (dly_flag!=-1)
            {
                show_dlv = true;
            }



            string SN = System_Info.KMS_key;
            if (SN!="")
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


                Console.WriteLine("Set the activation server...");
                p.StartInfo.Arguments = @"//B //Nologo " + Environment.SystemDirectory + @"\slmgr.vbs /skms " + Server_address;
                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                p.Start();
                p.WaitForExit();
                p.Close();

                Console.WriteLine("Activation start...");
                p.StartInfo.Arguments = @"//B //Nologo " + Environment.SystemDirectory + @"\slmgr.vbs /ato";
                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                p.Start();
                p.WaitForExit();
                p.Close();

                if (show_dlv)
                {
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
            else
            {
                Console.WriteLine("This Windows version no support!");
            }

        }
    }
}
