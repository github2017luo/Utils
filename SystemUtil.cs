using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PEGA.SI.One.Common
{
    public class SystemUtil
    {

        /// <summary>
        /// 输出CPU信息
        /// </summary>
        /// <returns></returns>
        public static int GetCPUInfo()
        {

            return  Convert.ToInt32(GetCPUCounter());
        }

        /// <summary>
        /// 获取CPU信息
        /// </summary>
        /// <returns></returns>
        private static object GetCPUCounter()
        {
            PerformanceCounter pc = new PerformanceCounter();
            pc.CategoryName = "Processor";
            pc.CounterName = "% Processor Time";
            pc.InstanceName = "_Total";
            dynamic Value_1 = pc.NextValue();
            System.Threading.Thread.Sleep(1000);
            dynamic Value_2 = pc.NextValue();
            return Value_2;
        }

        //定义内存的信息结构
        [StructLayout(LayoutKind.Sequential)]
        public struct MEMORY_INFO
        {
            public uint dwLength;
            public uint dwMemoryLoad;
            public uint dwTotalPhys;
            public uint dwAvailPhys;
            public uint dwTotalPageFile;
            public uint dwAvailPageFile;
            public uint dwTotalVirtual;
            public uint dwAvailVirtual;
        }


        [DllImport("kernel32")]
        private static extern void GetWindowsDirectory(StringBuilder WinDir, int count);

        [DllImport("kernel32")]
        private static extern void GetSystemDirectory(StringBuilder SysDir, int count);

        [DllImport("kernel32")]
        private static extern void GlobalMemoryStatus(ref MEMORY_INFO meminfo);

        public static float GetMemInfoP()
        {
            return (SystemUtil.GetHardDiskFreeSpace("C") / SystemUtil.GetHardDiskSpace())*100;
        }

        /// <summary>
        /// 获取内存信息
        /// </summary>
        /// <returns></returns>
        public static int GetMemInfo()
        {
            //调用GlobalMemoryStatus函数获取内存的相关信息
            MEMORY_INFO MemInfo = new MEMORY_INFO();
            GlobalMemoryStatus(ref MemInfo);
            //拼接字符串
            return (int)MemInfo.dwMemoryLoad;
        }

        /// <summary>
        /// 获取指定驱动器的空间总大小(单位为B) 
        /// 只需输入代表驱动器的字母即可 （大写） 
        /// </summary>
        /// <param name="str_HardDiskName"></param>
        /// <returns></returns>
        public static float GetHardDiskSpace(string str_HardDiskName)
        {
            float totalSize = new float();
            str_HardDiskName = str_HardDiskName + ":\\";
            System.IO.DriveInfo[] drives = System.IO.DriveInfo.GetDrives();
            foreach (System.IO.DriveInfo drive in drives)
            {
                if (drive.Name == str_HardDiskName)
                {
                    totalSize = drive.TotalSize / (1024 * 1024 * 1024);
                }
            }
            return totalSize;
        }

        public static float GetHardDiskSpace()
        {
            string str_HardDiskName = "C";
            float totalSize = new float();
            str_HardDiskName = str_HardDiskName + ":\\";
            System.IO.DriveInfo[] drives = System.IO.DriveInfo.GetDrives();
            foreach (System.IO.DriveInfo drive in drives)
            {
                if (drive.Name == str_HardDiskName)
                {
                    totalSize = drive.TotalSize / (1024 * 1024 * 1024);
                }
            }
            return totalSize;
        }

        /// <summary>
        /// 获取指定驱动器的剩余空间总大小(单位为B) 
        /// 只需输入代表驱动器的字母即可  
        /// </summary>
        /// <param name="str_HardDiskName"></param>
        /// <returns></returns>
        public static float GetHardDiskFreeSpace(string str_HardDiskName)
        {
            float freeSpace = new float();
            str_HardDiskName = str_HardDiskName + ":\\";
            System.IO.DriveInfo[] drives = System.IO.DriveInfo.GetDrives();
            foreach (System.IO.DriveInfo drive in drives)
            {
                if (drive.Name == str_HardDiskName)
                {
                    freeSpace = drive.TotalFreeSpace / (1024 * 1024 * 1024);
                }
            }
            return freeSpace;
        }
        public static int Get_TCP_Count()
        {
            IPGlobalProperties properties = IPGlobalProperties.GetIPGlobalProperties();
            TcpConnectionInformation[] connections = properties.GetActiveTcpConnections();
            return connections.Count();
        }
    }
}
