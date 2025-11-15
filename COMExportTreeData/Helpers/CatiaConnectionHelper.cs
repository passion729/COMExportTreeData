using System;
using System.Runtime.InteropServices;
using INFITF;

namespace COMExportTreeData.Helpers {
    /// <summary>
    /// CATIA连接辅助类
    /// </summary>
    public static class CatiaConnectionHelper {
        /// <summary>
        /// 连接到CATIA（若未运行则启动）
        /// </summary>
        public static Application ConnectToCatia() {
            Application catia;
            try {
                catia = (Application)Marshal.GetActiveObject("CATIA.Application");
            }
            catch {
                Type catiaType = Type.GetTypeFromProgID("CATIA.Application");
                catia = (Application)Activator.CreateInstance(catiaType);
                catia.Visible = false; // 批处理隐藏界面
            }

            return catia;
        }

        /// <summary>
        /// 释放CATIA资源
        /// </summary>
        public static void ReleaseCatia(Application catia) {
            if (catia != null) {
                Marshal.ReleaseComObject(catia);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}