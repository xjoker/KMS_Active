using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Management;

namespace KMS_Active
{
    static public class System_Info
    {
        private static string systemdrive = Environment.ExpandEnvironmentVariables("%systemdrive%");
        private static string Program_x86 = systemdrive+ "\\Program Files (x86)";
        private static string Program = systemdrive + "\\Program Files";

        static public string KMS_key
        {
            get
            {
                var scope = new ManagementScope("\\root\\cimv2");

                var query_Win32_OperatingSystem = new SelectQuery("Win32_OperatingSystem");
                var searcher_Win32_OperatingSyste = new ManagementObjectSearcher(scope, query_Win32_OperatingSystem);

                var query_Win32_ComputerSystem = new SelectQuery("Win32_ComputerSystem");
                var searcher_Win32_ComputerSystem = new ManagementObjectSearcher(scope, query_Win32_ComputerSystem);
                string SN = ""; //用于存储序列号作为返回值
                OperatingSystem osVersion = Environment.OSVersion;
                int majorVersion = osVersion.Version.Major;  //系统主版本
                int minorVersion = osVersion.Version.Minor;  //系统子版本
                int OperatingSystemSKU = -1;

                // 0=Unknow
                // 1=Workstation
                // 2=Domain Controller
                // 3=Server
                int PCSystemType = 0;



                foreach (var item in searcher_Win32_OperatingSyste.Get())
                {
                    if (item["Name"] != null)
                    {
                        //OperatingSystemSKU
                        OperatingSystemSKU = Convert.ToInt32(item["OperatingSystemSKU"]);
                    }
                }

                foreach (var item in searcher_Win32_ComputerSystem.Get())
                {
                    if (item["Name"] != null)
                    {
                        //PCSystemType
                        PCSystemType = Convert.ToInt32(item["PCSystemType"]);
                    }
                }

                if (majorVersion == 6)
                {
                    #region NT 6.0

                    switch (minorVersion)
                    {
                        case 0:
                            {
                                switch (PCSystemType)
                                {
                                    // Vista
                                    case 1:
                                        {
                                            switch (OperatingSystemSKU)
                                            {
                                                //  Windows Vista Enterprise
                                                case PRODUCT_ENTERPRISE_N:
                                                    SN = "VTC42-BM838-43QHV-84HX6-XJXKV";
                                                    break;
                                                case PRODUCT_ENTERPRISE:
                                                    SN = "VKK3X-68KWM-X2YGT-QR4M6-4BWMV";
                                                    break;

                                                // Windows Vista Business
                                                case PRODUCT_BUSINESS:
                                                    SN = "YFKBB-PQJJV-G996G-VWGXY-2V3X8";
                                                    break;
                                                case PRODUCT_BUSINESS_N:
                                                    SN = "HMBQG-8H2RH-C77VX-27R82-VMQBT";
                                                    break;
                                            }
                                        }
                                        break;
                                    // Server 2008
                                    case 3:
                                        {
                                            switch (OperatingSystemSKU)
                                            {
                                                // Windows Server 2008 Standard and Core
                                                case PRODUCT_STANDARD_SERVER:
                                                case PRODUCT_STANDARD_SERVER_CORE:
                                                    SN = "TM24T-X9RMF-VWXK6-X8JC9-BFGM2";
                                                    break;
                                                // Windows Server 2008 Standard without Hyper-V and Core
                                                case PRODUCT_STANDARD_SERVER_CORE_V:
                                                case PRODUCT_STANDARD_SERVER_V:
                                                    SN = "W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ";
                                                    break;
                                                // Windows Server 2008 Enterprise and Core
                                                case PRODUCT_ENTERPRISE_SERVER:
                                                case PRODUCT_ENTERPRISE_SERVER_CORE:
                                                    SN = "YQGMW-MPWTJ-34KDK-48M3W-X4Q6V";
                                                    break;
                                                // Windows Server 2008 Enterprise without Hyper-V and Core
                                                case PRODUCT_ENTERPRISE_SERVER_CORE_V:
                                                case PRODUCT_ENTERPRISE_SERVER_V:
                                                    SN = "39BXF-X8Q23-P2WWT-38T2F-G3FPG";
                                                    break;
                                                //  Windows Server 2008 HPC
                                                case PRODUCT_CLUSTER_SERVER:
                                                    SN = "RCTX3-KWVHP-BR6TB-RB6DM-6X7HP";
                                                    break;
                                                // Windows Server 2008 Datacenter and Core
                                                case PRODUCT_DATACENTER_SERVER:
                                                case PRODUCT_DATACENTER_SERVER_CORE:
                                                    SN = "7M67G-PC374-GR742-YH8V4-TCBY3";
                                                    break;
                                                //  Windows Server 2008 Datacenter without Hyper-V
                                                case PRODUCT_DATACENTER_SERVER_CORE_V:
                                                case PRODUCT_DATACENTER_SERVER_V:
                                                    SN = "22XQ2-VRXRG-P8D42-K34TD-G3QQC";
                                                    break;
                                            }

                                        }
                                        break;
                                }
                            }
                            break;
                        // windows 7
                        case 1:
                            {
                                switch (PCSystemType)
                                {
                                    case 1:
                                        {
                                            switch (OperatingSystemSKU)
                                            {
                                                //  Windows 7 Enterprise
                                                case PRODUCT_ENTERPRISE_N:
                                                    SN = "YDRBP-3D83W-TY26F-D46B2-XCKRJ";
                                                    break;
                                                case PRODUCT_ENTERPRISE:
                                                    SN = "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH";
                                                    break;

                                                // Windows 7 Professional
                                                case PRODUCT_BUSINESS:
                                                    SN = "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4";
                                                    break;
                                                case PRODUCT_BUSINESS_N:
                                                    SN = "MRPKT-YTG23-K7D7T-X2JMM-QY7MG";
                                                    break;
                                            }
                                        }
                                        break;
                                    case 3:
                                        {
                                            switch (OperatingSystemSKU)
                                            {
                                                // Windows Server 2008 R2 Web
                                                case PRODUCT_WEB_SERVER:
                                                case PRODUCT_WEB_SERVER_CORE:
                                                    SN = "6TPJF-RBVHG-WBW2R-86QPH-6RTM4";
                                                    break;
                                                // Windows Server 2008 R2 HPC edition
                                                case PRODUCT_CLUSTER_SERVER:
                                                    SN = "TT8MH-CG224-D3D7Q-498W2-9QCTX";
                                                    break;
                                                // Windows Server 2008 R2 Standard and Core
                                                case PRODUCT_STANDARD_SERVER:
                                                case PRODUCT_STANDARD_SERVER_CORE:
                                                    SN = "YC6KT-GKW9T-YTKYR-T4X34-R7VHC";
                                                    break;
                                                // Windows Server 2008 R2 Enterprise and Core
                                                case PRODUCT_ENTERPRISE_SERVER:
                                                case PRODUCT_ENTERPRISE_SERVER_CORE:
                                                    SN = "489J6-VHDMP-X63PK-3K798-CPX3Y";
                                                    break;
                                                // Windows Server 2008 R2 Datacenter and Core
                                                case PRODUCT_DATACENTER_SERVER:
                                                case PRODUCT_DATACENTER_SERVER_CORE:
                                                    SN = "74YFP-3QFB3-KQT8W-PMXWJ-7M648";
                                                    break;
                                            }
                                        }
                                        break;

                                }
                            }
                            break;
                        // windows 8
                        case 2:
                            {
                                switch (PCSystemType)
                                {
                                    case 1:
                                        {
                                            switch (OperatingSystemSKU)
                                            {
                                                //  Windows 8 Enterprise Edition
                                                case PRODUCT_ENTERPRISE_N:
                                                    SN = "JMNMF-RHW7P-DMY6X-RF3DR-X2BQT";
                                                    break;
                                                case PRODUCT_ENTERPRISE:
                                                    SN = "32JNW-9KQ84-P47T8-D8GGY-CWCK7";
                                                    break;

                                                // Windows 8 Professional Edition
                                                // Windows 8 Professional with Media Center Edition
                                                case PRODUCT_PROFESSIONAL_WMC:
                                                case PRODUCT_PROFESSIONAL:
                                                    SN = "NG4HW-VH26C-733KW-K6F98-J8CK4";
                                                    break;
                                                case PRODUCT_PROFESSIONAL_N:
                                                    SN = "XCVCF-2NXM9-723PB-MHCB7-2RYQQ";
                                                    break;
                                            }
                                        }
                                        break;
                                    case 3:
                                        {
                                            switch (OperatingSystemSKU)
                                            {
                                                // Windows Server 2012 Server Standard and Core
                                                case PRODUCT_STANDARD_SERVER:
                                                case PRODUCT_STANDARD_SERVER_CORE:
                                                    SN = "BN3D2-R7TKB-3YPBD-8DRP2-27GG4";
                                                    break;
                                                // Windows Server 2012 MultiPoint Standard
                                                case PRODUCT_MULTIPOINT_STANDARD_SERVER:
                                                    SN = "HM7DN-YVMH3-46JC3-XYTG7-CYQJJ";
                                                    break;
                                                // Windows Server 2012 MultiPoint Premium
                                                case PRODUCT_MULTIPOINT_PREMIUM_SERVER:
                                                    SN = "XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G";
                                                    break;
                                                // Windows Server 2012 Datacenter and Core
                                                case PRODUCT_DATACENTER_SERVER:
                                                case PRODUCT_DATACENTER_SERVER_CORE:
                                                    SN = "48HP8-DN98B-MYWDG-T2DCC-8W83P";
                                                    break;
                                                default:
                                                    SN = "48HP8-DN98B-MYWDG-T2DCC-8W83P";
                                                    break;
                                            }
                                        }
                                        break;

                                }
                            }
                            break;
                        // windows 8.1
                        case 3:
                            {
                                switch (PCSystemType)
                                {
                                    case 1:
                                        {
                                            switch (OperatingSystemSKU)
                                            {
                                                //  Windows 8.1 Enterprise Edition
                                                case PRODUCT_ENTERPRISE_N:
                                                    SN = "TT4HM-HN7YT-62K67-RGRQJ-JFFXW";
                                                    break;
                                                case PRODUCT_ENTERPRISE:
                                                    SN = "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7";
                                                    break;

                                                // Windows 8.1 Professional Edition
                                                // Windows 8.1 Professional with Media Center Edition
                                                case PRODUCT_PROFESSIONAL_WMC:
                                                case PRODUCT_PROFESSIONAL:
                                                    SN = "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9";
                                                    break;
                                                case PRODUCT_PROFESSIONAL_N:
                                                    SN = "HMCNV-VVBFX-7HMBH-CTY9B-B4FXY";
                                                    break;
                                            }
                                        }
                                        break;
                                    case 3:
                                        {
                                            switch (OperatingSystemSKU)
                                            {
                                                // Windows Server 2012 R2 Server Standard and Core
                                                case PRODUCT_STANDARD_SERVER:
                                                case PRODUCT_STANDARD_SERVER_CORE:
                                                    SN = "D2N9P-3P6X9-2R39C-7RTCD-MDVJX";
                                                    break;
                                                // Windows Server 2012 R2 Datacenter and Core
                                                case PRODUCT_DATACENTER_SERVER:
                                                case PRODUCT_DATACENTER_SERVER_CORE:
                                                    SN = "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9";
                                                    break;
                                                default:
                                                    SN = "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9";
                                                    break;
                                            }
                                        }
                                        break;
                                }
                            }
                            break;
                    }
                    #endregion
                }
                if (majorVersion==10)
                {
                    #region NT 10.0
                    switch (minorVersion)
                    {
                        case 0:
                            {
                                switch (PCSystemType)
                                {
                                    case 1:
                                        {
                                            switch (OperatingSystemSKU)
                                            {
                                                //  Windows 10 Enterprise
                                                case PRODUCT_ENTERPRISE_N:
                                                    SN = "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4";
                                                    break;
                                                case PRODUCT_ENTERPRISE:
                                                    SN = "NPPR9-FWDCX-D2C8J-H872K-2YT43";
                                                    break;

                                                // Windows 10 Professional
                                                case PRODUCT_PROFESSIONAL:
                                                    SN = "W269N-WFGWX-YVC9B-4J6C9-T83GX";
                                                    break;
                                                case PRODUCT_PROFESSIONAL_N:
                                                    SN = "MH37W-N47XK-V7XM9-C7227-GCQG9";
                                                    break;

                                                // Windows 10 Education
                                                case PRODUCT_EDUCATION:
                                                    SN = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2";
                                                    break;
                                                case PRODUCT_EDUCATION_N:
                                                    SN = "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ";
                                                    break;

                                                // Windows 10 Enterprise 2015 LTSB
                                                case PRODUCT_ENTERPRISE_S:
                                                    SN = "WNMTR-4C88C-JK8YV-HQ7T2-76DF9";
                                                    break;
                                                case PRODUCT_ENTERPRISE_S_N:
                                                    SN = "2F77B-TNFGY-69QQF-B8YKP-D69TJ";
                                                    break;


                                            }
                                        }
                                        break;
                                    case 3:
                                        {
                                            switch (OperatingSystemSKU)
                                            {
                                                // Windows Server 2016 Standard and Core
                                                case PRODUCT_STANDARD_SERVER:
                                                case PRODUCT_STANDARD_SERVER_CORE:
                                                    SN = "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY";
                                                    break;
                                                // Windows Server 2016 Datacenter and Core
                                                case PRODUCT_DATACENTER_SERVER:
                                                case PRODUCT_DATACENTER_SERVER_CORE:
                                                    SN = "CB7KF-BWN84-R7R2Y-793K2-8XDDG";
                                                    break;
                                                default:
                                                    SN = "CB7KF-BWN84-R7R2Y-793K2-8XDDG";
                                                    break;
                                            }
                                        }
                                        break;
                                }
                            }
                            break;

                    }
                    #endregion
                }
                
                return SN;
            }
        }

        static public string Office_KMS_key
        {
            get
            {
                string ospp = "";
                foreach (var path in office_install_path)
                {
                    if (System.IO.File.Exists(path))
                    {
                        ospp = path;
                        return ospp;
                    }
                }
                return ospp;
            }
        }


        #region 系统版本对应表
        // 对应表 https://msdn.microsoft.com/zh-cn/library/windows/desktop/ms724358(v=vs.85).aspx

        // server
        private const int PRODUCT_DATACENTER_EVALUATION_SERVER = 0x00000050;
        private const int PRODUCT_DATACENTER_SERVER = 0x00000008;
        private const int PRODUCT_DATACENTER_SERVER_CORE_V = 0x00000027;
        private const int PRODUCT_DATACENTER_SERVER_V = 0x00000025;
        private const int PRODUCT_CLUSTER_SERVER = 0x00000012;
        private const int PRODUCT_ENTERPRISE_SERVER_CORE_V = 0x00000029;
        private const int PRODUCT_ENTERPRISE_SERVER_V = 0x00000026;
        private const int PRODUCT_ENTERPRISE_SERVER = 0x0000000A;
        private const int PRODUCT_ENTERPRISE_SERVER_CORE = 0x0000000E;
        private const int PRODUCT_STANDARD_SERVER_CORE_V = 0x00000028;
        private const int PRODUCT_STANDARD_SERVER_V = 0x00000024;
        private const int PRODUCT_STANDARD_EVALUATION_SERVER = 0x0000004F;
        private const int PRODUCT_STANDARD_SERVER = 0x00000007;
        private const int PRODUCT_STANDARD_SERVER_CORE = 0x0000000D;
        private const int PRODUCT_WEB_SERVER = 0x00000011;
        private const int PRODUCT_WEB_SERVER_CORE = 0x0000001D;
        private const int PRODUCT_MULTIPOINT_PREMIUM_SERVER = 0x0000004D;
        private const int PRODUCT_MULTIPOINT_STANDARD_SERVER = 0x0000004C;
        private const int PRODUCT_SOLUTION_EMBEDDEDSERVER = 0x00000038;
        private const int PRODUCT_DATACENTER_SERVER_CORE = 0x0000000C;

        // not server
        private const int PRODUCT_UNDEFINED = 0x00000000;
        private const int PRODUCT_ULTIMATE = 0x00000001;
        private const int PRODUCT_HOME_BASIC = 0x00000002;
        private const int PRODUCT_HOME_PREMIUM = 0x00000003;
        private const int PRODUCT_ENTERPRISE = 0x00000004;
        private const int PRODUCT_ENTERPRISE_N = 0x0000001B;
        private const int PRODUCT_HOME_BASIC_N = 0x00000005;
        private const int PRODUCT_BUSINESS = 0x00000006;
        private const int PRODUCT_BUSINESS_N = 0x00000010;
        private const int PRODUCT_STARTER = 0x0000000B;
        private const int PRODUCT_PROFESSIONAL = 0x00000030;
        private const int PRODUCT_PROFESSIONAL_N = 0x00000031;
        private const int PRODUCT_PROFESSIONAL_WMC = 0x00000067;
        private const int PRODUCT_CORE = 0x00000065;
        private const int PRODUCT_CORE_N = 0x00000062;
        private const int PRODUCT_CORE_COUNTRYSPECIFIC = 0x00000063;
        private const int PRODUCT_MOBILE_CORE = 0x00000068;
        private const int PRODUCT_MOBILE_ENTERPRISE = 0x00000085;
        private const int PRODUCT_EDUCATION = 0x00000079;
        private const int PRODUCT_EDUCATION_N = 0x0000007A;
        private const int PRODUCT_ENTERPRISE_S = 0x0000007D;
        private const int PRODUCT_ENTERPRISE_S_EVALUATION = 0x00000081;
        private const int PRODUCT_ENTERPRISE_S_N = 0x0000007E;
        private const int PRODUCT_ENTERPRISE_S_N_EVALUATION = 0x00000082;

        #endregion

        #region Office OSPP path
        private static string[] office_install_path = new string[]
        {
            Program+@"\Microsoft Office\Office16\OSPP.VBS",
            Program_x86+@"\Microsoft Office\Office16\OSPP.VBS",
            Program+@"\Microsoft Office\Office15\OSPP.VBS",
            Program_x86+@"\Microsoft Office\Office15\OSPP.VBS",
            Program+@"\Microsoft Office\Office14\OSPP.VBS",
            Program_x86+@"\Microsoft Office\Office14\OSPP.VBS",
            Program+@"\Microsoft Office\Office12\OSPP.VBS",
            Program_x86+@"\Microsoft Office\Office12\OSPP.VBS",
        };
        #endregion
    }
}
