using ModbusZt;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using ReaLTaiizor.Colors;
using ReaLTaiizor.Controls;
using ReaLTaiizor.Enum.Material;
using ReaLTaiizor.Forms;
using ReaLTaiizor.Manager;
using ReaLTaiizor.Util;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Lifetime;
using System.Security.AccessControl;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace v2
{
    public partial class frmMain : Form
    {
        public static String version = "v20251220";

        private readonly MaterialSkinManager materialSkinManager;

        Thread m_threadRunVideoIn;
        public CUsbCamera cameraIn;
        public Boolean m_bVideoRunIn;
        public Object captureLockIn = new object();
        public List<Bitmap> m_ListBitmapIn;

        Thread m_threadRunVideoOut;
        public CUsbCamera cameraOut;
        public Boolean m_bVideoRunOut;
        public Object captureLockOut = new object();
        public List<Bitmap> m_ListBitmapOut;

        Thread m_threadRunVideoCfg;
        public CUsbCamera cameraCfg;
        public Boolean m_bVideoRunCfg;
        public Object captureLockCfg = new object();
        public List<Bitmap> m_ListBitmapCfg;

        Thread m_threadRunModbus;
        public Boolean m_bRunModbus;

        ModbusSlaveTCP modbusTcpSlave;
        // Create a Modbus database with one device
        Datastore[] modbusDB = new Datastore[]
        {
            new Datastore(1, 100, 100, 100, 100)  ////discrete_inputs, coils, input_registers, holding_registers 每个变量100个。
        };

        public CV2Cfg m_Cfg;
        public CV2Cfg m_CfgSetting;
        public Bitmap m_BitmapCfg;

        public List<CArea> m_CAreaListCur;
        public List<String> m_CAreaMatCur;

        public Boolean m_bBigArea;
        public Boolean m_bBeginArea;
        public CArea   m_cArea;


        public Boolean m_bVisualRun;

        public static frmMain m_self = null;

        public static bool m_bInCameraTestPause;
        public static bool m_bOutCameraTestPause;


        public frmMain()
        {
            m_self = this;
            InitializeComponent();

            HomeMenu.Side = NightPanel.PanelSide.Right;
            SettingMenu.Side = NightPanel.PanelSide.Left;

            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.EnforceBackcolorOnAllComponents = true;
            materialSkinManager.Theme = MaterialSkinManager.Themes.DARK;

            m_bVideoRunIn = true;
            m_ListBitmapIn = new List<Bitmap>();
            m_bVideoRunOut = true;
            m_ListBitmapOut = new List<Bitmap>();
            m_bVideoRunCfg = true;
            m_ListBitmapCfg = new List<Bitmap>();
            m_bVideoRunCfg = true;

            m_bRunModbus = true;

            modbusTcpSlave = new ModbusSlaveTCP(modbusDB, IPAddress.Any, 502);
            modbusTcpSlave.TCPClientConnected += (sender, e) =>
            {
                Console.WriteLine($"Client connected: {e.RemoteEndPoint}");
            };

            modbusTcpSlave.TCPClientDisconnected += (sender, e) =>
            {
                Console.WriteLine($"Client disconnected: {e.RemoteEndPoint}");
            };

            // Start listening for Modbus TCP connections
            modbusTcpSlave.StartListen();

            m_bVisualRun = false;

            m_bBeginArea = false;

            m_bInCameraTestPause = false;
            m_bOutCameraTestPause = false;

        }

        #region tools 

        public static bool fnIsRuning(Boolean bCreate)
        {
            Semaphore sem = null;
            String semaphoreName = "qq327689069v2";
            bool doesNotExist = false;
            bool unauthorized = false;

            // Attempt to open the named semaphore.
            try
            {
                sem = Semaphore.OpenExisting(semaphoreName);
                sem.Close();
                return true;
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                Console.WriteLine("Semaphore does not exist.");
                doesNotExist = true;
                if (bCreate == false) return false;
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine("Unauthorized access: {0}", ex.Message);
                unauthorized = true;
                if (bCreate == false) return false;
            }
            catch(Exception ee)
            {
                Console.WriteLine("Unauthorized access: {0}", ee.Message);
                if (bCreate == false) return false;
            }

            if (doesNotExist)
            {
                bool semaphoreWasCreated;

                string user = Environment.UserDomainName + "\\"
                    + Environment.UserName;
                SemaphoreSecurity semSec = new SemaphoreSecurity();

                SemaphoreAccessRule rule = new SemaphoreAccessRule(
                    user,
                    SemaphoreRights.Synchronize | SemaphoreRights.Modify,
                    AccessControlType.Deny);
                semSec.AddAccessRule(rule);

                rule = new SemaphoreAccessRule(
                    user,
                    SemaphoreRights.ReadPermissions | SemaphoreRights.ChangePermissions,
                    AccessControlType.Allow);
                semSec.AddAccessRule(rule);

                sem = new Semaphore(3, 3, semaphoreName,
                    out semaphoreWasCreated, semSec);
                sem.Close();

                if (semaphoreWasCreated)
                {
                    Console.WriteLine("Created the semaphore.");
                    return false;
                }
                else
                {
                    Logger.GetLogger("IsRunning").Error("Unable to create the semaphore.");
                }
            }
            else if (unauthorized)
            {
                try
                {
                    sem = Semaphore.OpenExisting(
                        semaphoreName,
                        SemaphoreRights.ReadPermissions
                            | SemaphoreRights.ChangePermissions);

                    SemaphoreSecurity semSec = sem.GetAccessControl();

                    string user = Environment.UserDomainName + "\\"
                        + Environment.UserName;

                    SemaphoreAccessRule rule = new SemaphoreAccessRule(
                        user,
                        SemaphoreRights.Synchronize | SemaphoreRights.Modify,
                        AccessControlType.Deny);
                    semSec.RemoveAccessRule(rule);

                    rule = new SemaphoreAccessRule(user,
                         SemaphoreRights.Synchronize | SemaphoreRights.Modify,
                         AccessControlType.Allow);
                    semSec.AddAccessRule(rule);

                    sem.SetAccessControl(semSec);

                    sem = Semaphore.OpenExisting(semaphoreName);

                    sem.Close();
                    return false;   //// 创建成功
                }
                catch (UnauthorizedAccessException ex)
                {
                    Logger.GetLogger("IsRunning").Error("Unable to change permissions: " + ex.Message);
                }
            }

            return true;
        }
        
        public static void fnSleep(int msec)
        {
            try
            {
                Thread.Sleep(msec);
            }
            catch (Exception ee)
            {
                Debug.Print(ee.Message);
            }

        }

        public static string MatToJson(Mat mat)
        {
            if (mat == null)
                throw new ArgumentNullException(nameof(mat));

            // 将 Mat 转换为 Bitmap
            Bitmap bitmap = BitmapConverter.ToBitmap(mat);

            // 将 Bitmap 转换为 Base64 字符串
            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                byte[] imageBytes = ms.ToArray();
                string base64String = Convert.ToBase64String(imageBytes);
                return base64String;
            }
        }

        public static Mat JsonToMat(string json)
        {
            if (string.IsNullOrEmpty(json))
                throw new ArgumentNullException(nameof(json));

            // 将 Base64 字符串转换为字节数组
            byte[] imageBytes = Convert.FromBase64String(json);

            // 将字节数组转换为 Bitmap
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                Bitmap bitmap = new Bitmap(ms);

                // 将 Bitmap 转换为 Mat
                Mat mat = BitmapConverter.ToMat(bitmap);
                return mat;
            }
        }

        public static Boolean fnVisualTest(Int32 inout, Int32 topbtm)
        {
            Boolean bRtn = false;

            if (m_self == null) return false;
            if (m_self.m_Cfg == null) return false;

            if (m_self.m_Cfg.m_bUseNewVisualTest)
            {
                bRtn = fnVisualTest2(inout, topbtm);
            }
            else
            {
                bRtn = fnVisualTest1(inout, topbtm);

            }
            return bRtn;
        }

        public static Boolean fnVisualTest1(Int32 inout, Int32 topbtm)  //// 进板，出板， 正面，反面
        {
            //// add by dream  未完成

            Mat matVideo = new Mat();
            Mat matFeature = new Mat();
            Mat dstImg = new Mat();

            Logger.GetLogger("fnVisualTest1").Info("进入视觉测试: " + inout.ToString() + "," + topbtm.ToString());


            if (m_self == null) return false;

            if( inout == 1)
            {
                if (frmMain.m_bInCameraTestPause == true)
                {
                    Logger.GetLogger("fnVisualTest1").Info("进板暂停识别状态，返回识别成功。");
                    m_self.BeginInvoke(new Action(() =>
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(DateTime.Now.ToString("HH:mm:ss ") + "收到进板视觉检测请求，暂停识别。 "));
                    }));
                    return true;
                }
            }
            if (inout == 2)
            {
                if (frmMain.m_bOutCameraTestPause == true)
                {
                    Logger.GetLogger("fnVisualTest1").Info("出板暂停识别状态，返回识别成功。");
                    m_self.BeginInvoke(new Action(() =>
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(DateTime.Now.ToString("HH:mm:ss ") + "收到进板视觉检测请求，暂停识别。 "));
                    }));
                    return true;
                }
            }

            //String curTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            String startTime, endTime, msg = "";

            startTime = DateTime.Now.ToString("HH:mm:ss ");
            m_self.BeginInvoke(new Action(() =>
            {
                if (inout == 1)
                {
                    if (topbtm == 1)
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(startTime + "进板正面-进行视觉检测... 3s "));
                    }
                    else if (topbtm == 2)
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(startTime + "进板反面-进行视觉检测... 3s "));
                    }
                }
                else if (inout == 2)
                {
                    if (topbtm == 1)
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(startTime + "出板正面-进行视觉检测... 3s "));
                    }
                    else if (topbtm == 2)
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(startTime + "出板反面-进行视觉检测... 3s "));
                    }
                }
            }));

            try
            {
                if (inout == 1)  //// 进板
                {
                    fnSleep(m_self.m_Cfg.m_nInCameraDelay * 1000);

                    if (topbtm == 1) //// 正面
                    {
                        msg = "进板正面";

                        lock (m_self.captureLockIn)
                        {
                            try
                            {
                                if (m_self.m_ListBitmapIn != null)
                                {
                                    Bitmap bmpc = null;
                                    if (m_self.m_ListBitmapIn.Count > 0)
                                    {
                                        bmpc = m_self.m_ListBitmapIn[m_self.m_ListBitmapIn.Count - 1];
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                    else
                                    {
                                        bmpc = m_self.cameraIn.fnCapture();
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                }
                                else
                                {
                                    Bitmap bmpc = m_self.cameraIn.fnCapture();
                                    matVideo = BitmapConverter.ToMat(bmpc);
                                }

                                int x, y, w, h;
                                x = m_self.m_Cfg.m_AreaIn.start.X;
                                y = m_self.m_Cfg.m_AreaIn.start.Y;
                                w = m_self.m_Cfg.m_AreaIn.end.X - x;
                                h = m_self.m_Cfg.m_AreaIn.end.Y - y;

                                matVideo = matVideo.Clone(new Rect(x, y, w, h));

                            }
                            catch (Exception ee)
                            {
                                Logger.GetLogger("fnVisualTest 读进板摄像头正面出错").Error(ee.Message);
                                matVideo = null;
                                msg += "，处理进板正面摄像头数据出错 - " + ee.Message;
                            }

                            if (matVideo != null)
                            {
                                //// 循环判断所有特征
                                bool bOk = false;
                                double minVal, maxVal;
                                OpenCvSharp.Point minLoc, maxLoc;

                                for (int ii = 0; ii < m_self.m_Cfg.m_CAreaMatInTop.Count; ii++)
                                {
                                    matFeature = JsonToMat(m_self.m_Cfg.m_CAreaMatInTop[ii]);
                                    dstImg = new Mat(matVideo.Rows - matFeature.Rows + 1, matVideo.Cols - matFeature.Cols + 1, MatType.CV_32F, (Scalar)1);
                                    Cv2.MatchTemplate(matVideo, matFeature, dstImg, TemplateMatchModes.CCoeffNormed);
                                    Cv2.MinMaxLoc(dstImg, out minVal, out maxVal, out minLoc, out maxLoc);

                                    if( maxVal > m_self.m_Cfg.m_fCvFac )
                                    {
                                        if (Math.Abs(m_self.m_Cfg.m_CAreaListInTop[ii].start.X - maxLoc.X) < 10 && Math.Abs(m_self.m_Cfg.m_CAreaListInTop[ii].start.Y - maxLoc.Y) < 10)
                                        {
                                            bOk = true;
                                        }
                                    }

                                    dstImg.Dispose();
                                    if (bOk == false)
                                    {
                                        //// 计算时间差的秒数，转换成字符串
                                        endTime = DateTime.Now.ToString("HH:mm:ss ");
                                        m_self.BeginInvoke(new Action(() =>
                                        {
                                            m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "进板正面视觉失败！"));
                                        }));

                                        String fileName = fnGetAppPath() + "/logs/ErrInTop" +
                                        DateTime.Now.Year.ToString() + "_" +
                                        DateTime.Now.Month.ToString() + "_" +
                                        DateTime.Now.Day.ToString() + "_" +
                                        DateTime.Now.Hour.ToString() + "_" +
                                        DateTime.Now.Minute.ToString() + "_" +
                                        DateTime.Now.Second.ToString();
                                        try
                                        {
                                            Bitmap bmp = matVideo.ToBitmap();
                                            if (bmp != null)
                                            {
                                                bmp.Save(fileName + "_0.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                            Bitmap bmpf = matFeature.ToBitmap();
                                            if (bmpf != null)
                                            {
                                                bmpf.Save(fileName + "_1.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                        }
                                        catch (Exception ee)
                                        {
                                            Logger.GetLogger("保存识别失败图片").Error("进板正面保存文件出错，" + ee.Message);
                                        }

                                        Logger.GetLogger("fnVisualTest1").Error("进板正面识别失败，请查看输出的图片。");

                                        matFeature.Dispose();
                                        matVideo.Dispose();
                                        return false;
                                    }
                                    matFeature.Dispose();
                                }

                                endTime = DateTime.Now.ToString("HH:mm:ss ");
                                m_self.BeginInvoke(new Action(() =>
                                {
                                    m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "进板正面视觉成功！"));
                                }));
                                Logger.GetLogger("fnVisualTest1").Info("进板正面识别成功。");

                                matVideo.Dispose();
                                return true;
                            }
                        }
                    }
                    else if (topbtm == 2)  //// 反面
                    {
                        msg = "进板反面";
                        lock (m_self.captureLockIn)
                        {
                            try
                            {
                                if (m_self.m_ListBitmapIn != null)
                                {
                                    Bitmap bmpc = null;
                                    if (m_self.m_ListBitmapIn.Count > 0)
                                    {
                                        bmpc = m_self.m_ListBitmapIn[m_self.m_ListBitmapIn.Count - 1];
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                    else
                                    {
                                        bmpc = m_self.cameraIn.fnCapture();
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                }
                                else
                                {
                                    Bitmap bmpc = m_self.cameraIn.fnCapture();
                                    matVideo = BitmapConverter.ToMat(bmpc);
                                }

                                int x, y, w, h;
                                x = m_self.m_Cfg.m_AreaIn.start.X;
                                y = m_self.m_Cfg.m_AreaIn.start.Y;
                                w = m_self.m_Cfg.m_AreaIn.end.X - x;
                                h = m_self.m_Cfg.m_AreaIn.end.Y - y;

                                matVideo = matVideo.Clone(new Rect(x, y, w, h));

                            }
                            catch (Exception ee)
                            {
                                Logger.GetLogger("fnVisualTest 读进板摄像头反面出错").Error(ee.Message);
                                matVideo = null;
                                msg += "，处理进板反面摄像头数据出错 - " + ee.Message;
                            }

                            if (matVideo != null)
                            {
                                //// 循环判断所有特征
                                bool bOk = false;
                                double minVal, maxVal;
                                OpenCvSharp.Point minLoc, maxLoc;

                                for (int ii = 0; ii < m_self.m_Cfg.m_CAreaMatInBtm.Count; ii++)
                                {
                                    matFeature = JsonToMat(m_self.m_Cfg.m_CAreaMatInBtm[ii]);
                                    dstImg = new Mat(matVideo.Rows - matFeature.Rows + 1, matVideo.Cols - matFeature.Cols + 1, MatType.CV_32F, (Scalar)1);
                                    Cv2.MatchTemplate(matVideo, matFeature, dstImg, TemplateMatchModes.CCoeffNormed);
                                    Cv2.MinMaxLoc(dstImg, out minVal, out maxVal, out minLoc, out maxLoc);

                                    if (maxVal > m_self.m_Cfg.m_fCvFac)
                                    {
                                        if (Math.Abs(m_self.m_Cfg.m_CAreaListInBtm[ii].start.X - maxLoc.X) < 10 && Math.Abs(m_self.m_Cfg.m_CAreaListInBtm[ii].start.Y - maxLoc.Y) < 10)
                                        {
                                            bOk = true;
                                        }
                                    }

                                    dstImg.Dispose();
                                    if (bOk == false)
                                    {
                                        endTime = DateTime.Now.ToString("HH:mm:ss ");
                                        m_self.BeginInvoke(new Action(() =>
                                        {
                                            m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "进板反面视觉失败！"));
                                        }));

                                        String fileName = fnGetAppPath() + "/logs/ErrInBtm" +
                                        DateTime.Now.Year.ToString() + "_" +
                                        DateTime.Now.Month.ToString() + "_" +
                                        DateTime.Now.Day.ToString() + "_" +
                                        DateTime.Now.Hour.ToString() + "_" +
                                        DateTime.Now.Minute.ToString() + "_" +
                                        DateTime.Now.Second.ToString();
                                        try
                                        {
                                            Bitmap bmp = matVideo.ToBitmap();
                                            if (bmp != null)
                                            {
                                                bmp.Save(fileName + "_0.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                            Bitmap bmpf = matFeature.ToBitmap();
                                            if (bmpf != null)
                                            {
                                                bmpf.Save(fileName + "_1.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                        }
                                        catch (Exception ee)
                                        {
                                            Logger.GetLogger("保存识别失败图片").Error("进板反面保存文件出错，" + ee.Message);
                                        }
                                        Logger.GetLogger("fnVisualTest1").Error("进板反面识别失败，请查看输出的图片。");

                                        matFeature.Dispose();
                                        matVideo.Dispose();
                                        return false;
                                    }
                                    matFeature.Dispose();
                                }

                                endTime = DateTime.Now.ToString("HH:mm:ss ");
                                m_self.BeginInvoke(new Action(() =>
                                {
                                    m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "进板反面视觉成功！"));
                                }));
                                Logger.GetLogger("fnVisualTest1").Info("进板反面识别成功。");

                                matVideo.Dispose();
                                return true;
                            }
                        }
                    }
                }
                else if (inout == 2) //// 出板
                {
                    fnSleep(m_self.m_Cfg.m_nOutCameraDelay * 1000);

                    if (topbtm == 1) //// 正面
                    {
                        msg = "出板正面";

                        lock (m_self.captureLockOut)
                        {
                            try
                            {
                                if (m_self.m_ListBitmapOut != null)
                                {
                                    Bitmap bmpc = null;
                                    if (m_self.m_ListBitmapOut.Count > 0)
                                    {
                                        bmpc = m_self.m_ListBitmapOut[m_self.m_ListBitmapOut.Count - 1];
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                    else
                                    {
                                        bmpc = m_self.cameraOut.fnCapture();
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }

                                    int x, y, w, h;
                                    x = m_self.m_Cfg.m_AreaOut.start.X;
                                    y = m_self.m_Cfg.m_AreaOut.start.Y;
                                    w = m_self.m_Cfg.m_AreaOut.end.X - x;
                                    h = m_self.m_Cfg.m_AreaOut.end.Y - y;

                                    matVideo = matVideo.Clone(new Rect(x, y, w, h));
                                }
                                else
                                {
                                    Bitmap bmpc = m_self.cameraOut.fnCapture();
                                    matVideo = BitmapConverter.ToMat(bmpc);
                                }
                            }
                            catch (Exception ee)
                            {
                                Logger.GetLogger("fnVisualTest 读出板摄像头正面出错").Error(ee.Message);
                                matVideo = null;
                                msg += "，处理出板正面摄像头数据出错 - " + ee.Message;
                            }

                            if (matVideo != null)
                            {
                                //// 循环判断所有特征
                                bool bOk = false;
                                double minVal, maxVal;
                                OpenCvSharp.Point minLoc, maxLoc;

                                for (int ii = 0; ii < m_self.m_Cfg.m_CAreaMatOutTop.Count; ii++)
                                {
                                    matFeature = JsonToMat(m_self.m_Cfg.m_CAreaMatOutTop[ii]);
                                    dstImg = new Mat(matVideo.Rows - matFeature.Rows + 1, matVideo.Cols - matFeature.Cols + 1, MatType.CV_32F, (Scalar)1);
                                    Cv2.MatchTemplate(matVideo, matFeature, dstImg, TemplateMatchModes.CCoeffNormed);
                                    Cv2.MinMaxLoc(dstImg, out minVal, out maxVal, out minLoc, out maxLoc);

                                    if (maxVal > m_self.m_Cfg.m_fCvFac)
                                    {
                                        if (Math.Abs(m_self.m_Cfg.m_CAreaListOutTop[ii].start.X - maxLoc.X) < 10 && Math.Abs(m_self.m_Cfg.m_CAreaListOutTop[ii].start.Y - maxLoc.Y) < 10)
                                        {
                                            bOk = true;
                                        }
                                    }

                                    dstImg.Dispose();
                                    if (bOk == false)
                                    {
                                        endTime = DateTime.Now.ToString("HH:mm:ss ");
                                        m_self.BeginInvoke(new Action(() =>
                                        {
                                            m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "出板正面视觉失败！"));
                                        }));

                                        String fileName = fnGetAppPath() + "/logs/ErrOutTop" +
                                        DateTime.Now.Year.ToString() + "_" +
                                        DateTime.Now.Month.ToString() + "_" +
                                        DateTime.Now.Day.ToString() + "_" +
                                        DateTime.Now.Hour.ToString() + "_" +
                                        DateTime.Now.Minute.ToString() + "_" +
                                        DateTime.Now.Second.ToString();
                                        try
                                        {
                                            Bitmap bmp = matVideo.ToBitmap();
                                            if (bmp != null)
                                            {
                                                bmp.Save(fileName + "_0.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                            Bitmap bmpf = matFeature.ToBitmap();
                                            if (bmpf != null)
                                            {
                                                bmpf.Save(fileName + "_1.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                        }
                                        catch (Exception ee)
                                        {
                                            Logger.GetLogger("保存识别失败图片").Info("出板正面保存文件出错，" + ee.Message);
                                        }
                                        Logger.GetLogger("fnVisualTest1").Info("出板正面识别失败，请查看输出的图片。");

                                        matVideo.Dispose();
                                        return false;
                                    }
                                    matFeature.Dispose();
                                }

                                endTime = DateTime.Now.ToString("HH:mm:ss ");
                                m_self.BeginInvoke(new Action(() =>
                                {
                                    m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "出板正面视觉成功！"));
                                }));
                                Logger.GetLogger("fnVisualTest1").Info("出板正面识别成功。");

                                matFeature.Dispose();
                                matVideo.Dispose();
                                return true;
                            }
                        }

                    }
                    else if (topbtm == 2)  //// 反面
                    {
                        msg = "出板反面";
                        lock (m_self.captureLockOut)
                        {
                            try
                            {
                                if (m_self.m_ListBitmapOut != null)
                                {
                                    Bitmap bmpc = null;
                                    if (m_self.m_ListBitmapOut.Count > 0)
                                    {
                                        bmpc = m_self.m_ListBitmapOut[m_self.m_ListBitmapOut.Count - 1];
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                    else
                                    {
                                        bmpc = m_self.cameraOut.fnCapture();
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                }
                                else
                                {
                                    Bitmap bmpc = m_self.cameraOut.fnCapture();
                                    matVideo = BitmapConverter.ToMat(bmpc);
                                }

                                int x, y, w, h;
                                x = m_self.m_Cfg.m_AreaOut.start.X;
                                y = m_self.m_Cfg.m_AreaOut.start.Y;
                                w = m_self.m_Cfg.m_AreaOut.end.X - x;
                                h = m_self.m_Cfg.m_AreaOut.end.Y - y;

                                matVideo = matVideo.Clone(new Rect(x, y, w, h));

                            }
                            catch (Exception ee)
                            {
                                Logger.GetLogger("fnVisualTest 读出板摄像头反面出错").Info(ee.Message);
                                matVideo = null;
                                msg += "，处理出板反面摄像头数据出错 - " + ee.Message;
                            }

                            if (matVideo != null)
                            {
                                //// 循环判断所有特征
                                bool bOk = false;
                                double minVal, maxVal;
                                OpenCvSharp.Point minLoc, maxLoc;

                                for (int ii = 0; ii < m_self.m_Cfg.m_CAreaMatOutBtm.Count; ii++)
                                {
                                    matFeature = JsonToMat(m_self.m_Cfg.m_CAreaMatOutBtm[ii]);
                                    dstImg = new Mat(matVideo.Rows - matFeature.Rows + 1, matVideo.Cols - matFeature.Cols + 1, MatType.CV_32F, (Scalar)1);
                                    Cv2.MatchTemplate(matVideo, matFeature, dstImg, TemplateMatchModes.CCoeffNormed);
                                    Cv2.MinMaxLoc(dstImg, out minVal, out maxVal, out minLoc, out maxLoc);

                                    if (maxVal > m_self.m_Cfg.m_fCvFac)
                                    {
                                        if (Math.Abs(m_self.m_Cfg.m_CAreaListOutBtm[ii].start.X - maxLoc.X) < 10 && Math.Abs(m_self.m_Cfg.m_CAreaListOutBtm[ii].start.Y - maxLoc.Y) < 10)
                                        {
                                            bOk = true;
                                        }
                                    }

                                    dstImg.Dispose();

                                    if (bOk == false)
                                    {
                                        endTime = DateTime.Now.ToString("HH:mm:ss ");
                                        m_self.BeginInvoke(new Action(() =>
                                        {
                                            m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "出板反面视觉失败！"));
                                        }));

                                        String fileName = fnGetAppPath() + "/logs/ErrOutBtm" +
                                        DateTime.Now.Year.ToString() + "_" +
                                        DateTime.Now.Month.ToString() + "_" +
                                        DateTime.Now.Day.ToString() + "_" +
                                        DateTime.Now.Hour.ToString() + "_" +
                                        DateTime.Now.Minute.ToString() + "_" +
                                        DateTime.Now.Second.ToString();
                                        try
                                        {
                                            Bitmap bmp = matVideo.ToBitmap();
                                            if (bmp != null)
                                            {
                                                bmp.Save(fileName + "_0.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                            Bitmap bmpf = matFeature.ToBitmap();
                                            if (bmpf != null)
                                            {
                                                bmpf.Save(fileName + "_1.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                        }
                                        catch (Exception ee)
                                        {
                                            Logger.GetLogger("保存识别失败图片").Error("出板反面保存文件出错，" + ee.Message);
                                        }
                                        Logger.GetLogger("fnVisualTest1").Error("出板反面识别失败，请查看输出的图片。");

                                        matFeature.Dispose();
                                        matVideo.Dispose();
                                        return false;
                                    }
                                    matFeature.Dispose();
                                }

                                endTime = DateTime.Now.ToString("HH:mm:ss ");
                                m_self.BeginInvoke(new Action(() =>
                                {
                                    m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "出板反面视觉成功！"));
                                }));
                                Logger.GetLogger("fnVisualTest1").Info("出板反面识别成功。");

                                matVideo.Dispose();
                                return bOk;
                            }
                        }
                    }
                }
            }
            catch(Exception ee)
            {
                Logger.GetLogger("视觉识别出错").Error(ee.Message);
                msg += "，视觉识别出错 - " + ee.Message;
            }
            endTime = DateTime.Now.ToString("HH:mm:ss ");
            m_self.BeginInvoke(new Action(() =>
            {
                m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + msg));
            }));

            return false;
        }

        public static Boolean fnVisualTest2(Int32 inout, Int32 topbtm)  //// 进板，出板， 正面，反面
        {
            //// add by dream  未完成

            Mat matVideo = new Mat();
            Mat matFeature = new Mat();
            Mat matFeatureArea = new Mat();
            Mat dstImg = new Mat();

            Logger.GetLogger("fnVisualTest2").Info("进入视觉测试二: " + inout.ToString() + "," + topbtm.ToString());

            if (m_self == null) return false;

            if (inout == 1)
            {
                if (frmMain.m_bInCameraTestPause == true)
                {
                    Logger.GetLogger("fnVisualTest2").Info("进板暂停识别状态，返回识别成功。");
                    m_self.BeginInvoke(new Action(() =>
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(DateTime.Now.ToString("HH:mm:ss ") + "收到进板视觉检测请求，暂停识别。 "));
                    }));
                    return true;
                }
            }

            if (inout == 2)
            {
                if (frmMain.m_bOutCameraTestPause == true)
                {
                    Logger.GetLogger("fnVisualTest1").Info("出板暂停识别状态，返回识别成功。");
                    m_self.BeginInvoke(new Action(() =>
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(DateTime.Now.ToString("HH:mm:ss ") + "收到进板视觉检测请求，暂停识别。 "));
                    }));
                    return true;
                }
            }

            //String curTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            String startTime, endTime, msg = "";

            startTime = DateTime.Now.ToString("HH:mm:ss ");
            m_self.BeginInvoke(new Action(() =>
            {
                if (inout == 1)
                {
                    if (topbtm == 1)
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(startTime + "进板正面-进行视觉检测... 3s "));
                    }
                    else if (topbtm == 2)
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(startTime + "进板反面-进行视觉检测... 3s "));
                    }
                }
                else if (inout == 2)
                {
                    if (topbtm == 1)
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(startTime + "出板正面-进行视觉检测... 3s "));
                    }
                    else if (topbtm == 2)
                    {
                        m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(startTime + "出板反面-进行视觉检测... 3s "));
                    }
                }
            }));

            try
            {
                if (inout == 1)  //// 进板
                {
                    fnSleep(m_self.m_Cfg.m_nInCameraDelay * 1000); //// 灯图像稳定

                    if (topbtm == 1) //// 正面
                    {
                        msg = "进板正面";

                        lock (m_self.captureLockIn)
                        {
                            try
                            {
                                if (m_self.m_ListBitmapIn != null)
                                {
                                    Bitmap bmpc = null;
                                    if (m_self.m_ListBitmapIn.Count > 0)
                                    {
                                        bmpc = m_self.m_ListBitmapIn[m_self.m_ListBitmapIn.Count - 1];
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                    else
                                    {
                                        bmpc = m_self.cameraIn.fnCapture();
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                }
                                else
                                {
                                    Bitmap bmpc = m_self.cameraIn.fnCapture();
                                    matVideo = BitmapConverter.ToMat(bmpc);
                                }

                            }
                            catch (Exception ee)
                            {
                                Logger.GetLogger("fnVisualTest 读进板摄像头正面出错").Error(ee.Message);
                                matVideo = null;
                                msg += "，处理进板正面摄像头数据出错 - " + ee.Message;
                            }

                            if (matVideo != null)
                            {
                                //// 循环判断所有特征
                                bool bOk = false;
                                double minVal, maxVal;
                                OpenCvSharp.Point minLoc, maxLoc;

                                for (int ii = 0; ii < m_self.m_Cfg.m_CAreaMatInTop.Count; ii++)
                                {
                                    matFeature = JsonToMat(m_self.m_Cfg.m_CAreaMatInTop[ii]);

                                    //// 先判断  m_self.m_Cfg.m_CAreaListInTop2[ii] 的区域是否存在于matVideo中
                                    if (m_self.m_Cfg.m_CAreaListInTop2[ii].start.X < 0 || m_self.m_Cfg.m_CAreaListInTop2[ii].start.Y < 0
                                        || m_self.m_Cfg.m_CAreaListInTop2[ii].end.X > matVideo.Cols
                                        || m_self.m_Cfg.m_CAreaListInTop2[ii].end.Y > matVideo.Rows)
                                    {
                                        msg += "，视觉识别区域超出图像范围。start("
                                            + m_self.m_Cfg.m_CAreaListInTop2[ii].start.X.ToString() + ", "
                                            + m_self.m_Cfg.m_CAreaListInTop2[ii].start.Y.ToString() + ") end("
                                            + m_self.m_Cfg.m_CAreaListInTop2[ii].end.X.ToString() + ", "
                                            + m_self.m_Cfg.m_CAreaListInTop2[ii].end.Y.ToString() + ") image("
                                            + matVideo.Cols.ToString() + ", " + matVideo.Rows.ToString() + ")";


                                        Logger.GetLogger("fnVisualTest2").Error("进板正面识别失败。" + msg);
                                        bOk = false;
                                        matFeatureArea = null;

                                    }
                                    else {
                                        //// 根据m_self.m_Cfg.m_CAreaListInTop2中的区域加上扩展长度 m_iFeatureAreaLenth，从matVideo中截取识别区域，用来进行特征匹配
                                        int fx, fy, fw, fh;
                                        fx = m_self.m_Cfg.m_CAreaListInTop2[ii].start.X - m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fx < 0) fx = 0;

                                        fy = m_self.m_Cfg.m_CAreaListInTop2[ii].start.Y - m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fy < 0) fy = 0;

                                        fw = m_self.m_Cfg.m_CAreaListInTop2[ii].end.X - m_self.m_Cfg.m_CAreaListInTop2[ii].start.X + 2 * m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fx + fw > matVideo.Cols) fw = matVideo.Cols - fx;

                                        fh = m_self.m_Cfg.m_CAreaListInTop2[ii].end.Y - m_self.m_Cfg.m_CAreaListInTop2[ii].start.Y + 2 * m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fy + fh > matVideo.Rows) fh = matVideo.Rows - fy;

                                        matFeatureArea = matVideo.Clone(new Rect(fx, fy, fw, fh));

                                        dstImg = new Mat(matFeatureArea.Rows - matFeature.Rows + 1, matFeatureArea.Cols - matFeature.Cols + 1, MatType.CV_32F, (Scalar)1);
                                        Cv2.MatchTemplate(matFeatureArea, matFeature, dstImg, TemplateMatchModes.CCoeffNormed);
                                        Cv2.MinMaxLoc(dstImg, out minVal, out maxVal, out minLoc, out maxLoc);

                                        if (maxVal > m_self.m_Cfg.m_fCvFac)
                                        {
                                            bOk = true;
                                        }
                                    }

                                    dstImg.Dispose();
                                    if (bOk == false)
                                    {
                                        //// 计算时间差的秒数，转换成字符串
                                        endTime = DateTime.Now.ToString("HH:mm:ss ");
                                        m_self.BeginInvoke(new Action(() =>
                                        {
                                            m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "进板正面视觉失败！"));
                                        }));

                                        String fileName = fnGetAppPath() + "/logs/ErrInTop" +
                                        DateTime.Now.Year.ToString() + "_" +
                                        DateTime.Now.Month.ToString() + "_" +
                                        DateTime.Now.Day.ToString() + "_" +
                                        DateTime.Now.Hour.ToString() + "_" +
                                        DateTime.Now.Minute.ToString() + "_" +
                                        DateTime.Now.Second.ToString();
                                        try
                                        {
                                            Bitmap bmp = matFeatureArea.ToBitmap();
                                            if (bmp != null)
                                            {
                                                bmp.Save(fileName + "_0.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                            Bitmap bmpf = matFeature.ToBitmap();
                                            if (bmpf != null)
                                            {
                                                bmpf.Save(fileName + "_1.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                        }
                                        catch (Exception ee)
                                        {
                                            Logger.GetLogger("保存识别失败图片").Error("进板正面保存文件出错，" + ee.Message);
                                        }

                                        Logger.GetLogger("fnVisualTest2").Error("进板正面识别失败，请查看输出的图片。");
                                        if (matFeatureArea != null) matFeatureArea.Dispose();
                                        if (matFeature != null) matFeature.Dispose();
                                        matVideo.Dispose();
                                        return false;
                                    }
                                    if (matFeatureArea != null) matFeatureArea.Dispose();
                                    if (matFeature != null) matFeature.Dispose();
                                }

                                endTime = DateTime.Now.ToString("HH:mm:ss ");
                                m_self.BeginInvoke(new Action(() =>
                                {
                                    m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "进板正面视觉成功！"));
                                }));
                                Logger.GetLogger("fnVisualTest2").Info("进板正面识别成功。");

                                if (matFeatureArea != null) matFeatureArea.Dispose();
                                if (matFeature != null) matFeature.Dispose();
                                matVideo.Dispose();
                                return true;
                            }
                        }
                    }
                    else if (topbtm == 2)  //// 反面
                    {
                        msg = "进板反面";
                        lock (m_self.captureLockIn)
                        {
                            try
                            {
                                if (m_self.m_ListBitmapIn != null)
                                {
                                    Bitmap bmpc = null;
                                    if (m_self.m_ListBitmapIn.Count > 0)
                                    {
                                        bmpc = m_self.m_ListBitmapIn[m_self.m_ListBitmapIn.Count - 1];
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                    else
                                    {
                                        bmpc = m_self.cameraIn.fnCapture();
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                }
                                else
                                {
                                    Bitmap bmpc = m_self.cameraIn.fnCapture();
                                    matVideo = BitmapConverter.ToMat(bmpc);
                                }

                            }
                            catch (Exception ee)
                            {
                                Logger.GetLogger("fnVisualTest 读进板摄像头反面出错").Error(ee.Message);
                                matVideo = null;
                                msg += "，处理进板反面摄像头数据出错 - " + ee.Message;
                            }

                            if (matVideo != null)
                            {
                                //// 循环判断所有特征
                                bool bOk = false;
                                double minVal, maxVal;
                                OpenCvSharp.Point minLoc, maxLoc;

                                for (int ii = 0; ii < m_self.m_Cfg.m_CAreaMatInBtm.Count; ii++)
                                {
                                    matFeature = JsonToMat(m_self.m_Cfg.m_CAreaMatInBtm[ii]);

                                    if (m_self.m_Cfg.m_CAreaListInBtm2[ii].start.X < 0 || m_self.m_Cfg.m_CAreaListInBtm2[ii].start.Y < 0
                                        || m_self.m_Cfg.m_CAreaListInBtm2[ii].end.X > matVideo.Cols
                                        || m_self.m_Cfg.m_CAreaListInBtm2[ii].end.Y > matVideo.Rows)
                                    {
                                        msg += "，视觉识别区域超出图像范围。start("
                                            + m_self.m_Cfg.m_CAreaListInBtm2[ii].start.X.ToString() + ", "
                                            + m_self.m_Cfg.m_CAreaListInBtm2[ii].start.Y.ToString() + ") end("
                                            + m_self.m_Cfg.m_CAreaListInBtm2[ii].end.X.ToString() + ", "
                                            + m_self.m_Cfg.m_CAreaListInBtm2[ii].end.Y.ToString() + ") image("
                                            + matVideo.Cols.ToString() + ", " + matVideo.Rows.ToString() + ")";


                                        Logger.GetLogger("fnVisualTest2").Error("进板反面识别失败。" + msg);
                                        bOk = false;
                                        matFeatureArea = null;
                                    }
                                    else
                                    {
                                        //// 根据m_self.m_Cfg.m_CAreaListInTop2中的区域加上扩展长度 m_iFeatureAreaLenth，从matVideo中截取识别区域，用来进行特征匹配
                                        int fx, fy, fw, fh;
                                        fx = m_self.m_Cfg.m_CAreaListInBtm2[ii].start.X - m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fx < 0) fx = 0;

                                        fy = m_self.m_Cfg.m_CAreaListInBtm2[ii].start.Y - m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fy < 0) fy = 0;

                                        fw = m_self.m_Cfg.m_CAreaListInBtm2[ii].end.X - m_self.m_Cfg.m_CAreaListInBtm2[ii].start.X + 2 * m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fx + fw > matVideo.Cols) fw = matVideo.Cols - fx;

                                        fh = m_self.m_Cfg.m_CAreaListInBtm2[ii].end.Y - m_self.m_Cfg.m_CAreaListInBtm2[ii].start.Y + 2 * m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fy + fh > matVideo.Rows) fh = matVideo.Rows - fy;

                                        matFeatureArea = matVideo.Clone(new Rect(fx, fy, fw, fh));

                                        dstImg = new Mat(matFeatureArea.Rows - matFeature.Rows + 1, matFeatureArea.Cols - matFeature.Cols + 1, MatType.CV_32F, (Scalar)1);
                                        Cv2.MatchTemplate(matFeatureArea, matFeature, dstImg, TemplateMatchModes.CCoeffNormed);
                                        Cv2.MinMaxLoc(dstImg, out minVal, out maxVal, out minLoc, out maxLoc);

                                        if (maxVal > m_self.m_Cfg.m_fCvFac)
                                        {
                                                bOk = true;
                                        }
                                    }

                                    dstImg.Dispose();
                                    if (bOk == false)
                                    {
                                        endTime = DateTime.Now.ToString("HH:mm:ss ");
                                        m_self.BeginInvoke(new Action(() =>
                                        {
                                            m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "进板反面视觉失败！"));
                                        }));

                                        String fileName = fnGetAppPath() + "/logs/ErrInBtm" +
                                        DateTime.Now.Year.ToString() + "_" +
                                        DateTime.Now.Month.ToString() + "_" +
                                        DateTime.Now.Day.ToString() + "_" +
                                        DateTime.Now.Hour.ToString() + "_" +
                                        DateTime.Now.Minute.ToString() + "_" +
                                        DateTime.Now.Second.ToString();
                                        try
                                        {
                                            Bitmap bmp = matFeatureArea.ToBitmap();
                                            if (bmp != null)
                                            {
                                                bmp.Save(fileName + "_0.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                            Bitmap bmpf = matFeature.ToBitmap();
                                            if (bmpf != null)
                                            {
                                                bmpf.Save(fileName + "_1.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                        }
                                        catch (Exception ee)
                                        {
                                            Logger.GetLogger("保存识别失败图片").Error("进板反面保存文件出错，" + ee.Message);
                                        }
                                        Logger.GetLogger("fnVisualTest1").Error("进板反面识别失败，请查看输出的图片。");

                                        if (matFeatureArea != null) matFeatureArea.Dispose();
                                        if (matFeature != null) matFeature.Dispose();
                                        matVideo.Dispose();
                                        return false;
                                    }
                                    if (matFeatureArea != null) matFeatureArea.Dispose();
                                    if (matFeature != null) matFeature.Dispose();
                                }

                                endTime = DateTime.Now.ToString("HH:mm:ss ");
                                m_self.BeginInvoke(new Action(() =>
                                {
                                    m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "进板反面视觉成功！"));
                                }));
                                Logger.GetLogger("fnVisualTest1").Info("进板反面识别成功。");

                                if (matFeatureArea != null) matFeatureArea.Dispose();
                                if (matFeature != null) matFeature.Dispose();
                                matVideo.Dispose();
                                return true;
                            }
                        }
                    }
                }
                else if (inout == 2) //// 出板
                {
                    fnSleep(m_self.m_Cfg.m_nOutCameraDelay * 1000);

                    if (topbtm == 1) //// 正面
                    {
                        msg = "出板正面";

                        lock (m_self.captureLockOut)
                        {
                            try
                            {
                                if (m_self.m_ListBitmapOut != null)
                                {
                                    Bitmap bmpc = null;
                                    if (m_self.m_ListBitmapOut.Count > 0)
                                    {
                                        bmpc = m_self.m_ListBitmapOut[m_self.m_ListBitmapOut.Count - 1];
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                    else
                                    {
                                        bmpc = m_self.cameraOut.fnCapture();
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                }
                                else
                                {
                                    Bitmap bmpc = m_self.cameraOut.fnCapture();
                                    matVideo = BitmapConverter.ToMat(bmpc);
                                }
                            }
                            catch (Exception ee)
                            {
                                Logger.GetLogger("fnVisualTest2 读出板摄像头正面出错").Error(ee.Message);
                                matVideo = null;
                                msg += "，处理出板正面摄像头数据出错 - " + ee.Message;
                            }

                            if (matVideo != null)
                            {
                                //// 循环判断所有特征
                                bool bOk = false;
                                double minVal, maxVal;
                                OpenCvSharp.Point minLoc, maxLoc;

                                for (int ii = 0; ii < m_self.m_Cfg.m_CAreaMatOutTop.Count; ii++)
                                {
                                    matFeature = JsonToMat(m_self.m_Cfg.m_CAreaMatOutTop[ii]);

                                    //// 先判断  m_self.m_Cfg.m_CAreaListInTop2[ii] 的区域是否存在于matVideo中
                                    if (m_self.m_Cfg.m_CAreaListOutTop2[ii].start.X < 0 || m_self.m_Cfg.m_CAreaListOutTop2[ii].start.Y < 0
                                        || m_self.m_Cfg.m_CAreaListOutTop2[ii].end.X > matVideo.Cols
                                        || m_self.m_Cfg.m_CAreaListOutTop2[ii].end.Y > matVideo.Rows)
                                    {
                                        msg += "，视觉识别区域超出图像范围。start("
                                            + m_self.m_Cfg.m_CAreaListOutTop2[ii].start.X.ToString() + ", "
                                            + m_self.m_Cfg.m_CAreaListOutTop2[ii].start.Y.ToString() + ") end("
                                            + m_self.m_Cfg.m_CAreaListOutTop2[ii].end.X.ToString() + ", "
                                            + m_self.m_Cfg.m_CAreaListOutTop2[ii].end.Y.ToString() + ") image("
                                            + matVideo.Cols.ToString() + ", " + matVideo.Rows.ToString() + ")";


                                        Logger.GetLogger("fnVisualTest2").Error("出板正面识别失败。" + msg);
                                        bOk = false;
                                        matFeatureArea = null;

                                    }
                                    else
                                    {
                                        //// 根据m_self.m_Cfg.m_CAreaListInTop2中的区域加上扩展长度 m_iFeatureAreaLenth，从matVideo中截取识别区域，用来进行特征匹配
                                        int fx, fy, fw, fh;
                                        fx = m_self.m_Cfg.m_CAreaListOutTop2[ii].start.X - m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fx < 0) fx = 0;

                                        fy = m_self.m_Cfg.m_CAreaListOutTop2[ii].start.Y - m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fy < 0) fy = 0;

                                        fw = m_self.m_Cfg.m_CAreaListOutTop2[ii].end.X - m_self.m_Cfg.m_CAreaListOutTop2[ii].start.X + 2 * m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fx + fw > matVideo.Cols) fw = matVideo.Cols - fx;

                                        fh = m_self.m_Cfg.m_CAreaListOutTop2[ii].end.Y - m_self.m_Cfg.m_CAreaListOutTop2[ii].start.Y + 2 * m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fy + fh > matVideo.Rows) fh = matVideo.Rows - fy;

                                        matFeatureArea = matVideo.Clone(new Rect(fx, fy, fw, fh));

                                        dstImg = new Mat(matFeatureArea.Rows - matFeature.Rows + 1, matFeatureArea.Cols - matFeature.Cols + 1, MatType.CV_32F, (Scalar)1);
                                        Cv2.MatchTemplate(matFeatureArea, matFeature, dstImg, TemplateMatchModes.CCoeffNormed);
                                        Cv2.MinMaxLoc(dstImg, out minVal, out maxVal, out minLoc, out maxLoc);

                                        if (maxVal > m_self.m_Cfg.m_fCvFac)
                                        {
                                                bOk = true;
                                        }
                                    }

                                    dstImg.Dispose();
                                    if (bOk == false)
                                    {
                                        endTime = DateTime.Now.ToString("HH:mm:ss ");
                                        m_self.BeginInvoke(new Action(() =>
                                        {
                                            m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "出板正面视觉失败！"));
                                        }));

                                        String fileName = fnGetAppPath() + "/logs/ErrOutTop" +
                                        DateTime.Now.Year.ToString() + "_" +
                                        DateTime.Now.Month.ToString() + "_" +
                                        DateTime.Now.Day.ToString() + "_" +
                                        DateTime.Now.Hour.ToString() + "_" +
                                        DateTime.Now.Minute.ToString() + "_" +
                                        DateTime.Now.Second.ToString();
                                        try
                                        {
                                            Bitmap bmp = matFeatureArea.ToBitmap();
                                            if (bmp != null)
                                            {
                                                bmp.Save(fileName + "_0.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                            Bitmap bmpf = matFeature.ToBitmap();
                                            if (bmpf != null)
                                            {
                                                bmpf.Save(fileName + "_1.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                        }
                                        catch (Exception ee)
                                        {
                                            Logger.GetLogger("保存识别失败图片").Info("出板正面保存文件出错，" + ee.Message);
                                        }
                                        Logger.GetLogger("fnVisualTest1").Info("出板正面识别失败，请查看输出的图片。");

                                        if (matFeatureArea != null) matFeatureArea.Dispose();
                                        if (matFeature != null) matFeature.Dispose();
                                        matVideo.Dispose();
                                        return false;
                                    }
                                    if (matFeatureArea != null) matFeatureArea.Dispose();
                                    if (matFeature != null) matFeature.Dispose();
                                }

                                endTime = DateTime.Now.ToString("HH:mm:ss ");
                                m_self.BeginInvoke(new Action(() =>
                                {
                                    m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "出板正面视觉成功！"));
                                }));
                                Logger.GetLogger("fnVisualTest1").Info("出板正面识别成功。");

                                if (matFeatureArea != null) matFeatureArea.Dispose();
                                if (matFeature != null) matFeature.Dispose();
                                matVideo.Dispose();
                                return true;
                            }
                        }

                    }
                    else if (topbtm == 2)  //// 反面
                    {
                        msg = "出板反面";
                        lock (m_self.captureLockOut)
                        {
                            try
                            {
                                if (m_self.m_ListBitmapOut != null)
                                {
                                    Bitmap bmpc = null;
                                    if (m_self.m_ListBitmapOut.Count > 0)
                                    {
                                        bmpc = m_self.m_ListBitmapOut[m_self.m_ListBitmapOut.Count - 1];
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                    else
                                    {
                                        bmpc = m_self.cameraOut.fnCapture();
                                        matVideo = BitmapConverter.ToMat(bmpc);
                                    }
                                }
                                else
                                {
                                    Bitmap bmpc = m_self.cameraOut.fnCapture();
                                    matVideo = BitmapConverter.ToMat(bmpc);
                                }
                            }
                            catch (Exception ee)
                            {
                                Logger.GetLogger("fnVisualTest 读出板摄像头反面出错").Info(ee.Message);
                                matVideo = null;
                                msg += "，处理出板反面摄像头数据出错 - " + ee.Message;
                            }

                            if (matVideo != null)
                            {
                                //// 循环判断所有特征
                                bool bOk = false;
                                double minVal, maxVal;
                                OpenCvSharp.Point minLoc, maxLoc;

                                for (int ii = 0; ii < m_self.m_Cfg.m_CAreaMatOutBtm.Count; ii++)
                                {

                                    matFeature = JsonToMat(m_self.m_Cfg.m_CAreaMatOutBtm[ii]);

                                    if (m_self.m_Cfg.m_CAreaListOutBtm2[ii].start.X < 0 || m_self.m_Cfg.m_CAreaListOutBtm2[ii].start.Y < 0
                                        || m_self.m_Cfg.m_CAreaListOutBtm2[ii].end.X > matVideo.Cols
                                        || m_self.m_Cfg.m_CAreaListOutBtm2[ii].end.Y > matVideo.Rows)
                                    {
                                        msg += "，视觉识别区域超出图像范围。start("
                                            + m_self.m_Cfg.m_CAreaListOutBtm2[ii].start.X.ToString() + ", "
                                            + m_self.m_Cfg.m_CAreaListOutBtm2[ii].start.Y.ToString() + ") end("
                                            + m_self.m_Cfg.m_CAreaListOutBtm2[ii].end.X.ToString() + ", "
                                            + m_self.m_Cfg.m_CAreaListOutBtm2[ii].end.Y.ToString() + ") image("
                                            + matVideo.Cols.ToString() + ", " + matVideo.Rows.ToString() + ")";


                                        Logger.GetLogger("fnVisualTest2").Error("出板反面识别失败。" + msg);
                                        bOk = false;
                                        matFeatureArea = null;
                                    }
                                    else
                                    {
                                        //// 根据m_self.m_Cfg.m_CAreaListInTop2中的区域加上扩展长度 m_iFeatureAreaLenth，从matVideo中截取识别区域，用来进行特征匹配
                                        int fx, fy, fw, fh;
                                        fx = m_self.m_Cfg.m_CAreaListOutBtm2[ii].start.X - m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fx < 0) fx = 0;

                                        fy = m_self.m_Cfg.m_CAreaListOutBtm2[ii].start.Y - m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fy < 0) fy = 0;

                                        fw = m_self.m_Cfg.m_CAreaListOutBtm2[ii].end.X - m_self.m_Cfg.m_CAreaListOutBtm2[ii].start.X + 2 * m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fx + fw > matVideo.Cols) fw = matVideo.Cols - fx;

                                        fh = m_self.m_Cfg.m_CAreaListOutBtm2[ii].end.Y - m_self.m_Cfg.m_CAreaListOutBtm2[ii].start.Y + 2 * m_self.m_Cfg.m_iFeatureAreaLenth;
                                        if (fy + fh > matVideo.Rows) fh = matVideo.Rows - fy;

                                        matFeatureArea = matVideo.Clone(new Rect(fx, fy, fw, fh));

                                        dstImg = new Mat(matFeatureArea.Rows - matFeature.Rows + 1, matFeatureArea.Cols - matFeature.Cols + 1, MatType.CV_32F, (Scalar)1);
                                        Cv2.MatchTemplate(matFeatureArea, matFeature, dstImg, TemplateMatchModes.CCoeffNormed);
                                        Cv2.MinMaxLoc(dstImg, out minVal, out maxVal, out minLoc, out maxLoc);

                                        if (maxVal > m_self.m_Cfg.m_fCvFac)
                                        {
                                                bOk = true;
                                        }

                                    }

                                    dstImg.Dispose();
                                    if (bOk == false)
                                    {
                                        endTime = DateTime.Now.ToString("HH:mm:ss ");
                                        m_self.BeginInvoke(new Action(() =>
                                        {
                                            m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "出板反面视觉失败！"));
                                        }));

                                        String fileName = fnGetAppPath() + "/logs/ErrOutBtm" +
                                        DateTime.Now.Year.ToString() + "_" +
                                        DateTime.Now.Month.ToString() + "_" +
                                        DateTime.Now.Day.ToString() + "_" +
                                        DateTime.Now.Hour.ToString() + "_" +
                                        DateTime.Now.Minute.ToString() + "_" +
                                        DateTime.Now.Second.ToString();
                                        try
                                        {
                                            Bitmap bmp = matVideo.ToBitmap();
                                            if (bmp != null)
                                            {
                                                bmp.Save(fileName + "_0.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                            Bitmap bmpf = matFeature.ToBitmap();
                                            if (bmpf != null)
                                            {
                                                bmpf.Save(fileName + "_1.png", System.Drawing.Imaging.ImageFormat.Png);
                                            }
                                        }
                                        catch (Exception ee)
                                        {
                                            Logger.GetLogger("保存识别失败图片").Error("出板反面保存文件出错，" + ee.Message);
                                        }
                                        Logger.GetLogger("fnVisualTest1").Error("出板反面识别失败，请查看输出的图片。");

                                        if (matFeatureArea != null) matFeatureArea.Dispose();
                                        if (matFeature != null) matFeature.Dispose();
                                        matVideo.Dispose();
                                        return false;
                                    }
                                    if (matFeatureArea != null) matFeatureArea.Dispose();
                                    if (matFeature != null) matFeature.Dispose();
                                }

                                endTime = DateTime.Now.ToString("HH:mm:ss ");
                                m_self.BeginInvoke(new Action(() =>
                                {
                                    m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + "出板反面视觉成功！"));
                                }));
                                Logger.GetLogger("fnVisualTest1").Info("出板反面识别成功。");

                                if (matFeatureArea != null) matFeatureArea.Dispose();
                                if (matFeature != null) matFeature.Dispose();
                                matVideo.Dispose();
                                return bOk;
                            }
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                Logger.GetLogger("视觉识别出错").Error(ee.Message);
                msg += "，视觉识别出错 - " + ee.Message;
            }
            endTime = DateTime.Now.ToString("HH:mm:ss ");
            m_self.BeginInvoke(new Action(() =>
            {
                m_self.materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(endTime + msg));
            }));

            return false;
        }

        public static String fnGetAppPath()
        {
            String strPath = Application.ExecutablePath;
            return Path.GetDirectoryName(strPath);// + Path.DirectorySeparatorChar;
        }

        public static CV2Cfg fnLoadConfig(String strCfgFile)
        {
            
            CV2Cfg cfg;
            String strBase64;
            JavaScriptSerializer js;

            if (!File.Exists(strCfgFile))
            {
                cfg = new CV2Cfg();
                fnSaveConfig(cfg, strCfgFile);
            }
            else
            {
                try
                {
                    js = new JavaScriptSerializer();
                    js.MaxJsonLength = Int32.MaxValue;

                    strBase64 = File.ReadAllText(strCfgFile, Encoding.UTF8);
                    cfg = js.Deserialize<CV2Cfg>(Encoding.Default.GetString(Convert.FromBase64String(strBase64)));
                }
                catch
                {
                    MessageBox.Show("文件错误！ " + strCfgFile);
                    cfg = null;
                }
            }
            return cfg;
        }

        public static void fnSaveConfig(CV2Cfg cfg)
        {
            String strBase64, strJson;
            JavaScriptSerializer js;

            String strCfgFile = fnGetAppPath() + "\\v2.cfg";

            if (cfg == null)
            {
                MessageBox.Show("配置文件为空！");
                return;
            }

            js = new JavaScriptSerializer();
            js.MaxJsonLength = Int32.MaxValue;

            strJson = js.Serialize(cfg);
            strBase64 = Convert.ToBase64String(Encoding.Default.GetBytes(strJson));

            FileStream fs = new FileStream(strCfgFile, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            sw.Write(strBase64);
            sw.Flush();
            sw.Close();
            fs.Close();
        }

        public static void fnSaveConfig(CV2Cfg cfg, String file)
        {
            String strBase64, strJson;
            JavaScriptSerializer js;

            String strCfgFile = file;

            if (cfg == null)
            {
                MessageBox.Show("配置为空！");
                return;
            }

            js = new JavaScriptSerializer();
            js.MaxJsonLength = Int32.MaxValue;

            strJson = js.Serialize(cfg);
            strBase64 = Convert.ToBase64String(Encoding.Default.GetBytes(strJson));

            FileStream fs = new FileStream(strCfgFile, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            sw.Write(strBase64);
            sw.Flush();
            sw.Close();
            fs.Close();
        }

        public String fnGetAreaMat( Bitmap source, CArea area )
        {
            Mat img = null;
            Mat src = BitmapConverter.ToMat(source);

            int x, y, w, h;
            x = area.start.X;
            y = area.start.Y;
            w = area.end.X - x;
            h = area.end.Y - y;

            img = src.Clone(new Rect(x, y, w, h));
            return MatToJson(img);
        }

        #endregion


        private void parrotPicBoxMenu_Click(object sender, EventArgs e)
        {
            nightPanelSidebar.Visible = !nightPanelSidebar.Visible;
        }

        private void parrotPicBoxMenu_MouseEnter(object sender, EventArgs e)
        {
            parrotPicBoxMenu.FilterEnabled = true;
        }

        private void parrotPicBoxMenu_MouseLeave(object sender, EventArgs e)
        {
            parrotPicBoxMenu.FilterEnabled = false;
        }

        private void HomeIcon_Click(object sender, EventArgs e)
        {
            cameraCfg.fnStopCamera();

            materialTabControlPages.SelectedTab = tabPageHome;

            HomeMenu.Side = NightPanel.PanelSide.Right;
            SettingMenu.Side = NightPanel.PanelSide.Left;
        }

        private void SettingMenu_Click(object sender, EventArgs e)
        {
            if( m_bVisualRun )
            {
                MessageBox.Show("正在提供视觉服务，请关闭服务后进行设置！");
                return;
            }

            cameraIn.fnStopCamera();
            cameraOut.fnStopCamera();

            materialTabControlPages.SelectedTab = tabPageSetting;

            SettingMenu.Side = NightPanel.PanelSide.Right;
            HomeMenu.Side = NightPanel.PanelSide.Left;
        }

        private void ExitMenu_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void materialButtonStartInCamera_Click(object sender, EventArgs e)
        {
            if (materialComboBoxInCamera.SelectedIndex < 0)
            {
                MessageBox.Show("选择进板摄像头！");
                return;
            }
            if (materialComboBoxInCameraPara.SelectedIndex < 0)
            {
                MessageBox.Show("选择进板摄像头格式！");
                return;
            }
            this.Cursor = Cursors.WaitCursor;
            if (cameraIn.fnStartCamera(materialComboBoxInCamera.SelectedIndex, materialComboBoxInCameraPara.SelectedIndex) == true)
            {
            }

            this.Cursor = Cursors.Default;

        }

        private void frmMain_Load(object sender, EventArgs e)
        {

            cameraIn = new CUsbCamera();
            cameraOut = new CUsbCamera();
            cameraCfg = new CUsbCamera();

            String[] cc = cameraIn.fnGetCamera();
            materialComboBoxInCamera.Items.Clear();
            materialComboBoxOutCamera.Items.Clear();
            materialComboBoxCfgCamera.Items.Clear();
            for (int i = 0; i < cc.Length; i++)
            {
                materialComboBoxInCamera.Items.Add(cc[i]);
                materialComboBoxOutCamera.Items.Add(cc[i]);
                materialComboBoxCfgCamera.Items.Add(cc[i]);
            }

            materialComboBoxInCamera.SelectedIndexChanged += (senderIn, args) =>
            {
                cc = cameraIn.fnGetForamt(materialComboBoxInCamera.SelectedIndex);
                materialComboBoxInCameraPara.Items.Clear();
                for (int i = 0; i < cc.Length; i++)
                {
                    materialComboBoxInCameraPara.Items.Add(cc[i]);
                }

            };
            materialComboBoxOutCamera.SelectedIndexChanged += (senderOut, args) =>
            {
                cc = cameraOut.fnGetForamt(materialComboBoxOutCamera.SelectedIndex);
                materialComboBoxOutCameraPara.Items.Clear();
                for (int i = 0; i < cc.Length; i++)
                {
                    materialComboBoxOutCameraPara.Items.Add(cc[i]);
                }

            };
            materialComboBoxCfgCamera.SelectedIndexChanged += (senderCfg, args) =>
            {
                cc = cameraCfg.fnGetForamt(materialComboBoxCfgCamera.SelectedIndex);
                materialComboBoxCfgCameraPara.Items.Clear();
                for (int i = 0; i < cc.Length; i++)
                {
                    materialComboBoxCfgCameraPara.Items.Add(cc[i]);
                }

            };

            timer1.Start();
            cameraIn.m_CameraName = "进板";
            cameraOut.m_CameraName = "出板";

            m_threadRunVideoIn = new Thread(fnThreadRunVideoIn);
            m_threadRunVideoIn.Start(this);

            m_threadRunVideoOut = new Thread(fnThreadRunVideoOut);
            m_threadRunVideoOut.Start(this);

            m_threadRunModbus = new Thread(fnThreadRunModbus);
            m_threadRunModbus.Start(this);

            m_threadRunVideoCfg = new Thread(fnThreadRunVideoCfg);
            m_threadRunVideoCfg.Start(this);

            //panelPicBox
            materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(DateTime.Now.ToString("HH:mm:ss ") + "启动视觉服务, 版本" + version));

        }

        public void fnThreadRunVideoIn(object msg)
        {
            frmMain fm = (frmMain)msg;

            while (fm.m_bVideoRunIn)
            {
                fnSleep(200);

                if (fm.Visible == true)
                {
                    lock (fm.captureLockIn)
                    {
                        Bitmap bmp = fm.cameraIn.fnCapture();

                        if (bmp != null)
                        {
                            Bitmap tmp;
                            for (int ii = 0; ii < fm.m_ListBitmapIn.Count; ii++)
                            {
                                tmp = fm.m_ListBitmapIn[ii];
                                tmp.Dispose();
                            }
                            fm.m_ListBitmapIn.Clear();
                            fm.m_ListBitmapIn.Add(bmp);
                        }
                    }
                }
            }
        }

        public void fnThreadRunVideoOut(object msg)
        {
            frmMain fm = (frmMain)msg;

            while (fm.m_bVideoRunOut)
            {
                fnSleep(200);

                if (fm.Visible == true)
                {
                    lock (fm.captureLockOut)
                    {
                        Bitmap bmp = fm.cameraOut.fnCapture();

                        if (bmp != null)
                        {
                            Bitmap tmp;
                            for (int ii = 0; ii < fm.m_ListBitmapOut.Count; ii++)
                            {
                                tmp = fm.m_ListBitmapOut[ii];
                                tmp.Dispose();
                            }
                            fm.m_ListBitmapOut.Clear();
                            fm.m_ListBitmapOut.Add(bmp);
                        }
                    }
                }
            }
        }

        public void fnThreadRunModbus(object msg)
        {
            frmMain fm = (frmMain)msg;

            while (fm.m_bRunModbus)
            {
                fnSleep(200);
            }
        }
        
        public void fnThreadRunVideoCfg(object msg)
        {
            frmMain fm = (frmMain)msg;

            while (fm.m_bVideoRunCfg)
            {
                fnSleep(200);

                if (fm.Visible == true)
                {
                    lock (fm.captureLockCfg)
                    {
                        Bitmap bmp = fm.cameraCfg.fnCapture();

                        if (bmp != null)
                        {
                            Bitmap tmp;
                            for (int ii = 0; ii < fm.m_ListBitmapCfg.Count; ii++)
                            {
                                tmp = fm.m_ListBitmapCfg[ii];
                                tmp.Dispose();
                            }
                            fm.m_ListBitmapCfg.Clear();
                            fm.m_ListBitmapCfg.Add(bmp);
                        }
                    }
                }
            }
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            if (this.Visible == true)
            {
                lock (captureLockIn)
                {
                    try
                    {
                        if (m_ListBitmapIn != null)
                        {
                            Bitmap bmpc = null;
                            if (m_ListBitmapIn.Count > 0)
                            {
                                bmpc = m_ListBitmapIn[m_ListBitmapIn.Count - 1];
                                m_ListBitmapIn.Remove(bmpc);
                            }

                            if (bmpc != null)
                            {
                                //Bitmap bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);
                                //Graphics dc = Graphics.FromImage(bmp);
                                //dc.DrawImage(bmpc, new Point(0, 0));

                                try
                                {
                                    Image bmp = pictureBoxInCamera.Image;
                                    pictureBoxInCamera.Image = bmpc;
                                    pictureBoxInCamera.Invalidate();
                                    if (bmp != null)
                                    {
                                        bmp.Dispose();
                                    }
                                }
                                catch (Exception ee)
                                {
                                    Debug.WriteLine("pictureBoxInCamera error 1 " + ee.Message);

                                }

                            }

                        }
                    }
                    catch (Exception ee)
                    {
                        Debug.WriteLine("pictureBoxInCamera error 2 " + ee.Message);
                    }
                }

                lock (captureLockOut)
                {
                    try
                    {
                        if (m_ListBitmapOut != null)
                        {
                            Bitmap bmpc = null;
                            if (m_ListBitmapOut.Count > 0)
                            {
                                bmpc = m_ListBitmapOut[m_ListBitmapOut.Count - 1];
                                m_ListBitmapOut.Remove(bmpc);
                            }

                            if (bmpc != null)
                            {
                                //Bitmap bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);
                                //Graphics dc = Graphics.FromImage(bmp);
                                //dc.DrawImage(bmpc, new Point(0, 0));

                                try
                                {
                                    Image bmp = pictureBoxOutCamera.Image;
                                    pictureBoxOutCamera.Image = bmpc;
                                    pictureBoxOutCamera.Invalidate();
                                    if (bmp != null)
                                    {
                                        bmp.Dispose();
                                    }
                                }
                                catch (Exception ee)
                                {
                                    Debug.WriteLine("pictureBoxOutCamera error 1 " + ee.Message);

                                }

                            }

                        }
                    }
                    catch (Exception ee)
                    {
                        Debug.WriteLine("pictureBoxOutCamera error 2 " + ee.Message);
                    }
                }

                lock (captureLockCfg)
                {
                    try
                    {
                        if (m_ListBitmapCfg != null)
                        {

                            if (m_ListBitmapCfg.Count > 0)
                            {
                                if(m_BitmapCfg != null ) m_BitmapCfg.Dispose();

                                m_BitmapCfg = m_ListBitmapCfg[m_ListBitmapCfg.Count - 1];
                                m_ListBitmapCfg.Remove(m_BitmapCfg);

                                pictureBoxCamera.Invalidate();
                            }
                        }
                    }
                    catch (Exception ee)
                    {
                        Debug.WriteLine("pictureBoxCfgCamera error 2 " + ee.Message);
                    }
                }
            }

            if(m_bVisualRun == true)
            {
                materialLabelVisual.Text = "视觉服务状态：运行";
                //materialLabelNet.Text = "网络服务状态：运行";
            }
            else
            {
                materialLabelVisual.Text = "视觉服务状态：停止";
                //materialLabelNet.Text = "网络服务状态：停止";
            }

        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            m_bVideoRunIn = false;
            m_bVideoRunOut = false;
            m_bVideoRunCfg = false;
            m_bRunModbus = false;

            modbusTcpSlave.StopListen();

        }

        private void materialButtonStartOutCamera_Click(object sender, EventArgs e)
        {
            if (materialComboBoxOutCamera.SelectedIndex < 0)
            {
                MessageBox.Show("选择出板摄像头！");
                return;
            }
            if (materialComboBoxOutCameraPara.SelectedIndex < 0)
            {
                MessageBox.Show("选择出板摄像头格式！");
                return;
            }
            this.Cursor = Cursors.WaitCursor;
            if (cameraOut.fnStartCamera(materialComboBoxOutCamera.SelectedIndex, materialComboBoxOutCameraPara.SelectedIndex) == true)
            {
            }

            this.Cursor = Cursors.Default;

        }

        private void materialButtonBrowseCfg1_Click(object sender, EventArgs e)
        {
            String strGCodeFile;

            OpenFileDialog dilog = new OpenFileDialog();
            dilog.Title = "请选择配置文件";
            dilog.Filter = "Config (*.cfg)|*.cfg";
            if (dilog.ShowDialog() == DialogResult.OK)
            {
                strGCodeFile = dilog.FileName;
                materialSingleLineTextFieldCfg.Text = strGCodeFile;
            }
            else
                materialSingleLineTextFieldCfg.Text = "";

        }

        private void materialButtonStart_Click(object sender, EventArgs e)
        {

            if (materialComboBoxInCamera.SelectedIndex < 0)
            {
                MessageBox.Show("选择进板摄像头！");
                return;
            }
            if (materialComboBoxInCameraPara.SelectedIndex < 0)
            {
                MessageBox.Show("选择进板摄像头格式！");
                return;
            }

            if (materialComboBoxOutCamera.SelectedIndex < 0)
            {
                MessageBox.Show("选择出板摄像头！");
                return;
            }
            if (materialComboBoxOutCameraPara.SelectedIndex < 0)
            {
                MessageBox.Show("选择出板摄像头格式！");
                return;
            }

            materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(DateTime.Now.ToString("HH:mm:ss ") + "启动进板摄像头 " + materialComboBoxInCamera.SelectedIndex + ", " + materialComboBoxInCameraPara.SelectedIndex));
            materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(DateTime.Now.ToString("HH:mm:ss ") + "启动出板摄像头 " + materialComboBoxOutCamera.SelectedIndex + ", " + materialComboBoxOutCameraPara.SelectedIndex));
            materialListBoxLog.Items.Insert(0, new ReaLTaiizor.Child.Material.MaterialListBoxItem(DateTime.Now.ToString("HH:mm:ss ") + "参数：" + m_Cfg.fnPara2String2() ));
            Logger.GetLogger("启动视觉服务").Info("启动视觉服务, 版本" + version);
            Logger.GetLogger("启动视觉服务").Info("识别参数：" + m_Cfg.fnPara2String());

            this.Cursor = Cursors.WaitCursor;

            cameraIn.fnStopCamera();
            cameraIn.fnStartCamera(materialComboBoxInCamera.SelectedIndex, materialComboBoxInCameraPara.SelectedIndex);

            cameraOut.fnStopCamera();
            cameraOut.fnStartCamera(materialComboBoxOutCamera.SelectedIndex, materialComboBoxOutCameraPara.SelectedIndex);

            m_bVisualRun = true;
            this.Cursor = Cursors.Default;

        }

        private void materialButtonStop_Click(object sender, EventArgs e)
        {
            cameraIn.fnStopCamera();
            cameraOut.fnStopCamera();
            m_bVisualRun = false;
        }

        private void materialButtonStopInCamera_Click(object sender, EventArgs e)
        {
            cameraIn.fnStopCamera();
        }

        private void materialButtonStopOutCamera_Click(object sender, EventArgs e)
        {
            cameraOut.fnStopCamera();
        }

        private void materialButtonLoadConfig_Click(object sender, EventArgs e)
        {
            //strCfgFile = fnGetAppPath() + "\\v2.cfg";

            if (String.IsNullOrEmpty(materialSingleLineTextFieldCfg.Text))
            {
                MessageBox.Show("请选择配置文件！");
                return;
            }
            m_Cfg = fnLoadConfig(materialSingleLineTextFieldCfg.Text);

            if (m_Cfg == null)
            {
                return;
            }

            materialComboBoxInCamera.SelectedIndex = m_Cfg.m_nInCamera;
            materialComboBoxInCameraPara.SelectedIndex = m_Cfg.m_nInCameraPara;
            materialComboBoxOutCamera.SelectedIndex = m_Cfg.m_nOutCamera;
            materialComboBoxOutCameraPara.SelectedIndex = m_Cfg.m_nOutCameraPara;

        }



        #region Setting 

        private void materialButtonCfgBrowse_Click(object sender, EventArgs e)
        {
            String strGCodeFile;

            OpenFileDialog dilog = new OpenFileDialog();
            dilog.CheckFileExists = false;
            dilog.Title = "请选择配置文件";
            dilog.Filter = "Config (*.cfg)|*.cfg";
            if (dilog.ShowDialog() == DialogResult.OK)
            {
                strGCodeFile = dilog.FileName;
                materialTextBoxEditCfgFile.Text = strGCodeFile;
            }
            else
                materialTextBoxEditCfgFile.Text = "";

        }



        private void materialButtonVideoCfgStart_Click(object sender, EventArgs e)
        {
            //Debug.WriteLine(metroPanelPicBox.AutoScroll);
            if (materialComboBoxCfgCamera.SelectedIndex < 0)
            {
                MessageBox.Show("选择摄像头！");
                return;
            }
            if (materialComboBoxCfgCameraPara.SelectedIndex < 0)
            {
                MessageBox.Show("选择摄像头格式！");
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            if (cameraCfg.fnStartCamera(materialComboBoxCfgCamera.SelectedIndex, materialComboBoxCfgCameraPara.SelectedIndex) == true)
            {
            }
            this.Cursor = Cursors.Default;

        }

        private void materialButtonVideoCfgStop_Click(object sender, EventArgs e)
        {
            cameraCfg.fnStopCamera();

            if (String.IsNullOrEmpty(materialTextBoxEditCfgFile.Text))
            {
                MessageBox.Show("请选择配置文件！");
                return;
            }

            if (materialComboBoxInOut.Text.Equals("进板摄像头"))
            {
                if (MessageBox.Show("是否保存当前摄像头为 进板摄像头？", "确定进板摄像头", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    m_CfgSetting.m_nInCamera = materialComboBoxCfgCamera.SelectedIndex;
                    m_CfgSetting.m_nInCameraPara = materialComboBoxCfgCameraPara.SelectedIndex;
                    fnSaveConfig(m_CfgSetting, materialTextBoxEditCfgFile.Text);
                }
            }
            else
            {
                if (MessageBox.Show("是否保存当前摄像头为 出板摄像头？", "确定出板摄像头", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    m_CfgSetting.m_nOutCamera = materialComboBoxCfgCamera.SelectedIndex;
                    m_CfgSetting.m_nOutCameraPara = materialComboBoxCfgCameraPara.SelectedIndex;
                    fnSaveConfig(m_CfgSetting, materialTextBoxEditCfgFile.Text);
                }
            }
        }

        private void materialButtonLoadCfg_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(materialTextBoxEditCfgFile.Text))
            {
                MessageBox.Show("请选择配置文件！");
                return;
            }
            m_CfgSetting = fnLoadConfig(materialTextBoxEditCfgFile.Text);

            if (m_CfgSetting == null)
            {
                return;
            }

            if(materialComboBoxInOut.Text.Equals("进板摄像头"))
            {

                if(materialComboBoxTopBtm.Text.Equals("板子正面"))
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListInTop;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatInTop;
                }
                else
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListInBtm;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatInBtm;
                }
            }
            else
            {
                if (materialComboBoxTopBtm.Text.Equals("板子正面"))
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListOutTop;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatOutTop;
                }
                else
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListOutBtm;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatOutBtm;
                }
            }

            if (m_CAreaListCur != null)
            {
                if (m_CAreaListCur.Count > 0)
                {
                    materialListBoxAreaList.Items.Clear();
                    for (int ii = 0; ii < m_CAreaListCur.Count; ii++)
                    {
                        materialListBoxAreaList.Items.Add( new ReaLTaiizor.Child.Material.MaterialListBoxItem(m_CAreaListCur[ii].ToString()) );
                    }
                }
            }
            else 
            {
                m_CAreaListCur = new List<CArea>();
                m_CAreaMatCur = new List<String>();
            }

            materialTextBoxEditFat.Text = m_CfgSetting.m_fCvFac.ToString();
            materialTextBoxEditInCamDelay.Text = m_CfgSetting.m_nInCameraDelay.ToString();
            materialTextBoxEditOutCamDelay.Text = m_CfgSetting.m_nOutCameraDelay.ToString();

            checkBoxNewVisualTest.Checked = m_CfgSetting.m_bUseNewVisualTest;
        }

        private void materialButtonSaveCfg_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(materialTextBoxEditCfgFile.Text))
            {
                MessageBox.Show("请选择配置文件！");
                return;
            }
            if(m_CfgSetting != null)
            {
                try
                {
                    m_CfgSetting.m_fCvFac = Double.Parse(materialTextBoxEditFat.Text);
                }
                catch( Exception ee)
                {
                    MessageBox.Show("系数错误，" + ee.Message);
                    m_CfgSetting.m_fCvFac = 0.8;
                }

                try
                {
                    m_CfgSetting.m_nInCameraDelay = Int32.Parse(materialTextBoxEditInCamDelay.Text);
                }
                catch (Exception ee)
                {
                    MessageBox.Show("进板延迟错误，" + ee.Message);
                    m_CfgSetting.m_nInCameraDelay = 3;
                }

                try
                {
                    m_CfgSetting.m_nOutCameraDelay = Int32.Parse(materialTextBoxEditOutCamDelay.Text);
                }
                catch (Exception ee)
                {
                    MessageBox.Show("出板延迟错误，" + ee.Message);
                    m_CfgSetting.m_nOutCameraDelay = 3;
                }

                if( m_CfgSetting.fnTestConfig() == false)
                {
                    return;
                }

                m_CfgSetting.m_bUseNewVisualTest = checkBoxNewVisualTest.Checked;

                fnSaveConfig(m_CfgSetting, materialTextBoxEditCfgFile.Text);
            }

        }

        private void materialButtonAreaAdd_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(materialTextBoxEditCfgFile.Text))
            {
                MessageBox.Show("请选择配置文件！");
                return;
            }

            if (materialComboBoxInOut.Text.Equals("进板摄像头"))
            {
                if (m_CfgSetting.m_AreaIn.start.X == 0 && m_CfgSetting.m_AreaIn.start.Y == 0)
                {
                    MessageBox.Show("请先选择一个进板识别范围！");
                    return;
                }
            }
            else
            {
                if (m_CfgSetting.m_AreaOut.start.X == 0 && m_CfgSetting.m_AreaOut.start.Y == 0)
                {
                    MessageBox.Show("请先选择一个出板识别范围！");
                    return;
                }
            }

            cameraCfg.fnStopCamera();

            this.Cursor = Cursors.Cross;
            m_bBeginArea = true;
            m_cArea = new CArea();

        }

        private void materialButtonAreaDel_Click(object sender, EventArgs e)
        {
            if (materialListBoxAreaList.SelectedIndex == -1)
            {
                MessageBox.Show("请选择一个识别区域！");
                return;
            }

            m_CAreaListCur.RemoveAt(materialListBoxAreaList.SelectedIndex);
            m_CAreaMatCur.RemoveAt(materialListBoxAreaList.SelectedIndex);
            materialListBoxAreaList.Items.RemoveAt(materialListBoxAreaList.SelectedIndex);

            pictureBoxCamera.Invalidate();

        }
        
        public void fnAddAreaMat(CArea area)
        {

            if (materialComboBoxInOut.Text.Equals("进板摄像头"))
            {
                if (materialComboBoxTopBtm.Text.Equals("板子正面"))
                {
                    m_CfgSetting.m_CAreaListInTop2.Add(area);
                    m_CfgSetting.m_CAreaMatInTop.Add(fnGetAreaMat(m_BitmapCfg, area));

                    CArea aa = new CArea();
                    aa.start.X = area.start.X - m_CfgSetting.m_AreaIn.start.X;
                    aa.start.Y = area.start.Y - m_CfgSetting.m_AreaIn.start.Y;
                    aa.end.X = area.end.X - m_CfgSetting.m_AreaIn.start.X;
                    aa.end.Y = area.end.Y - m_CfgSetting.m_AreaIn.start.Y;
                    m_CfgSetting.m_CAreaListInTop.Add(aa);
                    materialListBoxAreaList.Items.Add(new ReaLTaiizor.Child.Material.MaterialListBoxItem(aa.ToString()));

                }
                else
                {
                    m_CfgSetting.m_CAreaListInBtm2.Add(area);
                    m_CfgSetting.m_CAreaMatInBtm.Add(fnGetAreaMat(m_BitmapCfg, area));

                    CArea aa = new CArea();
                    aa.start.X = area.start.X - m_CfgSetting.m_AreaIn.start.X;
                    aa.start.Y = area.start.Y - m_CfgSetting.m_AreaIn.start.Y;
                    aa.end.X = area.end.X - m_CfgSetting.m_AreaIn.start.X;
                    aa.end.Y = area.end.Y - m_CfgSetting.m_AreaIn.start.Y;
                    m_CfgSetting.m_CAreaListInBtm.Add(aa);
                    materialListBoxAreaList.Items.Add(new ReaLTaiizor.Child.Material.MaterialListBoxItem(aa.ToString()));

                }
            }
            else
            {
                if (materialComboBoxTopBtm.Text.Equals("板子正面"))
                {
                    m_CfgSetting.m_CAreaListOutTop2.Add(area);
                    m_CfgSetting.m_CAreaMatOutTop.Add(fnGetAreaMat(m_BitmapCfg, area));

                    CArea aa = new CArea();
                    aa.start.X = area.start.X - m_CfgSetting.m_AreaOut.start.X;
                    aa.start.Y = area.start.Y - m_CfgSetting.m_AreaOut.start.Y;
                    aa.end.X = area.end.X - m_CfgSetting.m_AreaOut.start.X;
                    aa.end.Y = area.end.Y - m_CfgSetting.m_AreaOut.start.Y;
                    m_CfgSetting.m_CAreaListOutTop.Add(aa);
                    materialListBoxAreaList.Items.Add(new ReaLTaiizor.Child.Material.MaterialListBoxItem(aa.ToString()));
                }
                else
                {
                    m_CfgSetting.m_CAreaListOutBtm2.Add(area);
                    m_CfgSetting.m_CAreaMatOutBtm.Add(fnGetAreaMat(m_BitmapCfg, area));

                    CArea aa = new CArea();
                    aa.start.X = area.start.X - m_CfgSetting.m_AreaOut.start.X;
                    aa.start.Y = area.start.Y - m_CfgSetting.m_AreaOut.start.Y;
                    aa.end.X = area.end.X - m_CfgSetting.m_AreaOut.start.X;
                    aa.end.Y = area.end.Y - m_CfgSetting.m_AreaOut.start.Y;
                    m_CfgSetting.m_CAreaListOutBtm.Add(aa);
                    materialListBoxAreaList.Items.Add(new ReaLTaiizor.Child.Material.MaterialListBoxItem(aa.ToString()));
                }
            }
        }

        private void pictureBoxCamera_MouseClick(object sender, MouseEventArgs e)
        {
            if(m_bBeginArea == true)
            {
                if (m_cArea.start.X == 0 && m_cArea.start.Y == 0)
                {
                    m_cArea.start.X = e.X;
                    m_cArea.start.Y = e.Y;

                    this.Cursor = Cursors.Hand;
                }
                else
                {
                    m_cArea.end.X = e.X;
                    m_cArea.end.Y = e.Y;

                    CArea tmp = new CArea();
                    tmp.start.X = m_cArea.start.X < m_cArea.end.X ? m_cArea.start.X : m_cArea.end.X;
                    tmp.start.Y = m_cArea.start.Y < m_cArea.end.Y ? m_cArea.start.Y : m_cArea.end.Y;
                    tmp.end.X = m_cArea.start.X > m_cArea.end.X ? m_cArea.start.X : m_cArea.end.X;
                    tmp.end.Y = m_cArea.start.Y > m_cArea.end.Y ? m_cArea.start.Y : m_cArea.end.Y;
                    fnAddAreaMat(tmp);

                    m_cArea.start.X = 0;
                    m_cArea.start.Y = 0;
                    m_cArea.end.X = 0;
                    m_cArea.end.Y = 0;

                    this.Cursor = Cursors.Default;

                    m_bBeginArea = false;
                }
            }

            if( m_bBigArea == true )
            {
                if (materialComboBoxInOut.Text.Equals("进板摄像头"))
                {
                    if (m_CfgSetting.m_AreaIn.start.X == 0 && m_CfgSetting.m_AreaIn.start.Y == 0)
                    {
                        m_CfgSetting.m_AreaIn.start.X = e.X;
                        m_CfgSetting.m_AreaIn.start.Y = e.Y;
                        m_CfgSetting.m_AreaIn.end.X = 0;
                        m_CfgSetting.m_AreaIn.end.Y = 0;

                        this.Cursor = Cursors.Hand;
                    }
                    else
                    {

                        CArea tmp = new CArea();
                        tmp.start.X = m_CfgSetting.m_AreaIn.start.X < e.X ? m_CfgSetting.m_AreaIn.start.X : e.X;
                        tmp.start.Y = m_CfgSetting.m_AreaIn.start.Y < e.Y ? m_CfgSetting.m_AreaIn.start.Y : e.Y;
                        tmp.end.X = m_CfgSetting.m_AreaIn.start.X > e.X ? m_CfgSetting.m_AreaIn.start.X : e.X;
                        tmp.end.Y = m_CfgSetting.m_AreaIn.start.Y > e.Y ? m_CfgSetting.m_AreaIn.start.Y : e.Y;

                        m_CfgSetting.m_AreaIn.start.X = tmp.start.X;
                        m_CfgSetting.m_AreaIn.start.Y = tmp.start.Y;
                        m_CfgSetting.m_AreaIn.end.X = tmp.end.X;
                        m_CfgSetting.m_AreaIn.end.Y = tmp.end.Y;

                        fnSaveConfig(m_CfgSetting, materialTextBoxEditCfgFile.Text);

                        this.Cursor = Cursors.Default;

                        pictureBoxCamera.Invalidate();
                        m_bBigArea = false;
                    }
                }
                else
                {
                    if (m_CfgSetting.m_AreaOut.start.X == 0 && m_CfgSetting.m_AreaOut.start.Y == 0)
                    {
                        m_CfgSetting.m_AreaOut.start.X = e.X;
                        m_CfgSetting.m_AreaOut.start.Y = e.Y;
                        m_CfgSetting.m_AreaOut.end.X = 0;
                        m_CfgSetting.m_AreaOut.end.Y = 0;

                        this.Cursor = Cursors.Hand;
                    }
                    else
                    {

                        CArea tmp = new CArea();
                        tmp.start.X = m_CfgSetting.m_AreaOut.start.X < e.X ? m_CfgSetting.m_AreaOut.start.X : e.X;
                        tmp.start.Y = m_CfgSetting.m_AreaOut.start.Y < e.Y ? m_CfgSetting.m_AreaOut.start.Y : e.Y;
                        tmp.end.X = m_CfgSetting.m_AreaOut.start.X > e.X ? m_CfgSetting.m_AreaOut.start.X : e.X;
                        tmp.end.Y = m_CfgSetting.m_AreaOut.start.Y > e.Y ? m_CfgSetting.m_AreaOut.start.Y : e.Y;

                        m_CfgSetting.m_AreaOut.start.X = tmp.start.X;
                        m_CfgSetting.m_AreaOut.start.Y = tmp.start.Y;
                        m_CfgSetting.m_AreaOut.end.X = tmp.end.X;
                        m_CfgSetting.m_AreaOut.end.Y = tmp.end.Y;

                        fnSaveConfig(m_CfgSetting, materialTextBoxEditCfgFile.Text);

                        this.Cursor = Cursors.Default;

                        pictureBoxCamera.Invalidate();
                        m_bBigArea = false;
                    }
                }

            }
        }

        private void pictureBoxCamera_MouseMove(object sender, MouseEventArgs e)
        {
            if (m_bBeginArea == true)
            {
                if (m_cArea.start.X != 0 || m_cArea.start.Y != 0)
                {
                    m_cArea.end.X = e.X;
                    m_cArea.end.Y = e.Y;

                    pictureBoxCamera.Invalidate();
                }
            }
            if (m_bBigArea == true)
            {
                if (materialComboBoxInOut.Text.Equals("进板摄像头"))
                {
                    if (m_CfgSetting.m_AreaIn.start.X != 0 || m_CfgSetting.m_AreaIn.start.Y != 0)
                    {
                        m_CfgSetting.m_AreaIn.end.X = e.X;
                        m_CfgSetting.m_AreaIn.end.Y = e.Y;

                        pictureBoxCamera.Invalidate();
                    }
                }
                else
                {
                    if (m_CfgSetting.m_AreaOut.start.X != 0 || m_CfgSetting.m_AreaOut.start.Y != 0)
                    {
                        m_CfgSetting.m_AreaOut.end.X = e.X;
                        m_CfgSetting.m_AreaOut.end.Y = e.Y;

                        pictureBoxCamera.Invalidate();
                    }
                }
            }

            materialTextBoxEditXY.Text = e.X.ToString() + ", " + e.Y.ToString();
            if(m_CfgSetting != null)
            {
                if (materialComboBoxInOut.Text.Equals("进板摄像头"))
                {
                    materialTextBoxEditXYInner.Text = (e.X - m_CfgSetting.m_AreaIn.start.X).ToString() + ", " + (e.Y - m_CfgSetting.m_AreaIn.start.Y).ToString();
                }
                else
                {
                    materialTextBoxEditXYInner.Text = (e.X - m_CfgSetting.m_AreaIn.start.X).ToString() + ", " + (e.Y - m_CfgSetting.m_AreaOut.start.Y).ToString();
                }
            }

        }

        private void pictureBoxCamera_Paint(object sender, PaintEventArgs e)
        {
            if (m_BitmapCfg != null)
            {
                Bitmap imageCache;

                panelPic.Width = m_BitmapCfg.Width;
                panelPic.Height = m_BitmapCfg.Height;

                pictureBoxCamera.Width = m_BitmapCfg.Width;
                pictureBoxCamera.Height = m_BitmapCfg.Height;

                imageCache = new Bitmap(m_BitmapCfg.Width, m_BitmapCfg.Height);

                Rectangle des = new Rectangle(0, 0, m_BitmapCfg.Width, m_BitmapCfg.Height);
                Graphics dc = Graphics.FromImage(imageCache);
                dc.DrawImage(m_BitmapCfg, des);

                //// 画识别区域
                if (materialComboBoxInOut.Text.Equals("进板摄像头"))
                {

                    if (m_CfgSetting != null && m_CfgSetting.m_AreaIn.start.X != 0 && m_CfgSetting.m_AreaIn.start.Y != 0 &&
                        m_CfgSetting.m_AreaIn.end.X != 0 && m_CfgSetting.m_AreaIn.end.Y != 0)
                    {
                        int x, y, h, w;
                        x = m_CfgSetting.m_AreaIn.start.X;
                        y = m_CfgSetting.m_AreaIn.start.Y;
                        w = Math.Abs(x - m_CfgSetting.m_AreaIn.end.X);
                        h = Math.Abs(y - m_CfgSetting.m_AreaIn.end.Y);
                        dc.DrawRectangle(new Pen(Color.Blue, 4), x, y, w, h);
                    }
                }
                else
                {
                    if (m_CfgSetting != null && m_CfgSetting.m_AreaOut.start.X != 0 && m_CfgSetting.m_AreaOut.start.Y != 0 &&
                        m_CfgSetting.m_AreaOut.end.X != 0 && m_CfgSetting.m_AreaOut.end.Y != 0)
                    {
                        int x, y, h, w;
                        x = m_CfgSetting.m_AreaOut.start.X;
                        y = m_CfgSetting.m_AreaOut.start.Y;
                        w = Math.Abs(x - m_CfgSetting.m_AreaOut.end.X);
                        h = Math.Abs(y - m_CfgSetting.m_AreaOut.end.Y);
                        dc.DrawRectangle(new Pen(Color.Blue, 4), x, y, w, h);
                    }
                }

                //// 画已经有的区域，选中的区域用黄色，其他为蓝色
                if (m_CAreaListCur != null)
                {
                    int jj = materialListBoxAreaList.SelectedIndex;
                    for (int ii = 0; ii < m_CAreaListCur.Count; ii++ )
                    {
                        int x, y, h, w;
                        if (materialComboBoxInOut.Text.Equals("进板摄像头"))
                        {
                            x = m_CAreaListCur[ii].start.X + m_CfgSetting.m_AreaIn.start.X;
                            y = m_CAreaListCur[ii].start.Y + m_CfgSetting.m_AreaIn.start.Y;
                            w = m_CAreaListCur[ii].end.X - x + m_CfgSetting.m_AreaIn.start.X;
                            h = m_CAreaListCur[ii].end.Y - y + m_CfgSetting.m_AreaIn.start.Y;
                        }
                        else
                        {
                            x = m_CAreaListCur[ii].start.X + m_CfgSetting.m_AreaOut.start.X;
                            y = m_CAreaListCur[ii].start.Y + m_CfgSetting.m_AreaOut.start.Y;
                            w = m_CAreaListCur[ii].end.X - x + m_CfgSetting.m_AreaOut.start.X;
                            h = m_CAreaListCur[ii].end.Y - y + m_CfgSetting.m_AreaOut.start.Y;
                        }

                        if ( jj == ii )
                        {
                            dc.DrawRectangle(new Pen(Color.Yellow, 2), x, y, w, h);
                        }
                        else
                        {
                            dc.DrawRectangle(new Pen(Color.Blue, 2), x, y, w, h);
                        }
                    }
                }

                //// 画当前区域
                if (m_bBeginArea)
                {
                    if (m_cArea.start.X != 0 || m_cArea.start.Y != 0)
                    {
                        int x, y, h, w;
                        x = m_cArea.start.X < m_cArea.end.X ? m_cArea.start.X : m_cArea.end.X;
                        y = m_cArea.start.Y < m_cArea.end.Y ? m_cArea.start.Y : m_cArea.end.Y;
                        w = Math.Abs(m_cArea.start.X - m_cArea.end.X);
                        h = Math.Abs(m_cArea.start.Y - m_cArea.end.Y);
                        dc.DrawRectangle(new Pen(Color.Red, 2), x, y, w, h);
                    }
                }

                //// 画当前识别区域
                if (m_bBigArea)
                {
                    if (materialComboBoxInOut.Text.Equals("进板摄像头"))
                    {
                        if (m_CfgSetting.m_AreaIn.start.X != 0 || m_CfgSetting.m_AreaIn.start.Y != 0)
                        {
                            int x, y, h, w;

                            x = m_CfgSetting.m_AreaIn.start.X < m_CfgSetting.m_AreaIn.end.X ? m_CfgSetting.m_AreaIn.start.X : m_CfgSetting.m_AreaIn.end.X;
                            y = m_CfgSetting.m_AreaIn.start.Y < m_CfgSetting.m_AreaIn.end.Y ? m_CfgSetting.m_AreaIn.start.Y : m_CfgSetting.m_AreaIn.end.Y;
                            w = Math.Abs(m_CfgSetting.m_AreaIn.start.X - m_CfgSetting.m_AreaIn.end.X);
                            h = Math.Abs(m_CfgSetting.m_AreaIn.start.Y - m_CfgSetting.m_AreaIn.end.Y);
                            dc.DrawRectangle(new Pen(Color.Blue, 4), x, y, w, h);
                        }

                    }
                    else
                    {
                        if (m_CfgSetting.m_AreaOut.start.X != 0 || m_CfgSetting.m_AreaOut.start.Y != 0)
                        {
                            int x, y, h, w;

                            x = m_CfgSetting.m_AreaOut.start.X < m_CfgSetting.m_AreaOut.end.X ? m_CfgSetting.m_AreaOut.start.X : m_CfgSetting.m_AreaOut.end.X;
                            y = m_CfgSetting.m_AreaOut.start.Y < m_CfgSetting.m_AreaOut.end.Y ? m_CfgSetting.m_AreaOut.start.Y : m_CfgSetting.m_AreaOut.end.Y;
                            w = Math.Abs(m_CfgSetting.m_AreaOut.start.X - m_CfgSetting.m_AreaOut.end.X);
                            h = Math.Abs(m_CfgSetting.m_AreaOut.start.Y - m_CfgSetting.m_AreaOut.end.Y);
                            dc.DrawRectangle(new Pen(Color.Blue, 4), x, y, w, h);
                        }

                    }

                }

                e.Graphics.DrawImage(imageCache, des);
                dc.Dispose();
                imageCache.Dispose();
            }
        }


        #endregion

        private void materialComboBoxInOut_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (materialComboBoxInOut.Text.Equals("进板摄像头"))
            {

                if (materialComboBoxTopBtm.Text.Equals("板子正面"))
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListInTop;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatInTop;
                }
                else
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListInBtm;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatInBtm;
                }
            }
            else
            {

                if (materialComboBoxTopBtm.Text.Equals("板子正面"))
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListOutTop;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatOutTop;
                }
                else
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListOutBtm;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatOutBtm;
                }
            }

            if (m_CAreaListCur != null)
            {
                if (m_CAreaListCur.Count > 0)
                {
                    materialListBoxAreaList.Items.Clear();
                    for (int ii = 0; ii < m_CAreaListCur.Count; ii++)
                    {
                        materialListBoxAreaList.Items.Add(new ReaLTaiizor.Child.Material.MaterialListBoxItem(m_CAreaListCur[ii].ToString()));
                    }
                }
                else
                {
                    materialListBoxAreaList.Items.Clear();
                    m_CAreaListCur = new List<CArea>();
                    m_CAreaMatCur = new List<String>();
                }
            }
            else
            {
                materialListBoxAreaList.Items.Clear();
                m_CAreaListCur = new List<CArea>();
                m_CAreaMatCur = new List<String>();
            }

            pictureBoxCamera.Invalidate();
        }

        private void materialComboBoxTopBtm_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (materialComboBoxInOut.Text.Equals("进板摄像头"))
            {

                if (materialComboBoxTopBtm.Text.Equals("板子正面"))
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListInTop;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatInTop;
                }
                else
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListInBtm;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatInBtm;
                }
            }
            else
            {

                if (materialComboBoxTopBtm.Text.Equals("板子正面"))
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListOutTop;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatOutTop;
                }
                else
                {
                    m_CAreaListCur = m_CfgSetting.m_CAreaListOutBtm;
                    m_CAreaMatCur = m_CfgSetting.m_CAreaMatOutBtm;
                }
            }

            if (m_CAreaListCur != null)
            {
                if (m_CAreaListCur.Count > 0)
                {
                    materialListBoxAreaList.Items.Clear();
                    for (int ii = 0; ii < m_CAreaListCur.Count; ii++)
                    {
                        materialListBoxAreaList.Items.Add(new ReaLTaiizor.Child.Material.MaterialListBoxItem(m_CAreaListCur[ii].ToString()));
                    }
                }
                else
                {
                    materialListBoxAreaList.Items.Clear();
                    m_CAreaListCur = new List<CArea>();
                    m_CAreaMatCur = new List<String>();
                }
            }
            else
            {
                materialListBoxAreaList.Items.Clear();
                m_CAreaListCur = new List<CArea>();
                m_CAreaMatCur = new List<String>();
            }
            pictureBoxCamera.Invalidate();
        }

        private void materialButtonCapture_Click(object sender, EventArgs e)
        {
            String fileName = fnGetAppPath() + "/img" +
                DateTime.Now.Year.ToString() + "_" +
                DateTime.Now.Month.ToString() + "_" +
                DateTime.Now.Day.ToString() + "_" +
                DateTime.Now.Hour.ToString() + "_" +
                DateTime.Now.Minute.ToString() + "_" +
                DateTime.Now.Second.ToString() + ".png";


            lock (captureLockCfg)
            {
                try
                {
                    Bitmap bmp = cameraCfg.fnCapture();
                    if (bmp != null)
                    {
                        bmp.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
                        MessageBox.Show("图片存储在 " + fileName);
                    }
                    else
                    {
                        MessageBox.Show("摄像头不在预览状态！");
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.Message);
                }
            }

        }

        private void materialButtonArea_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(materialTextBoxEditCfgFile.Text))
            {
                MessageBox.Show("请选择配置文件！");
                return;
            }

            cameraCfg.fnStopCamera();

            this.Cursor = Cursors.Cross;
            m_bBigArea = true;
            if (materialComboBoxInOut.Text.Equals("进板摄像头"))
            {
                m_CfgSetting.m_AreaIn.start.X = 0;
                m_CfgSetting.m_AreaIn.start.Y = 0;
                m_CfgSetting.m_AreaIn.end.X = 0;
                m_CfgSetting.m_AreaIn.end.Y = 0;
            }
            else
            {
                m_CfgSetting.m_AreaOut.start.X = 0;
                m_CfgSetting.m_AreaOut.start.Y = 0;
                m_CfgSetting.m_AreaOut.end.X = 0;
                m_CfgSetting.m_AreaOut.end.Y = 0;
            }
            pictureBoxCamera.Invalidate();
        }

        private void materialListBoxAreaList_SelectedIndexChanged(object sender, ReaLTaiizor.Child.Material.MaterialListBoxItem selectedItem)
        {
            pictureBoxCamera.Invalidate();

        }

        private void materialCheckBoxInCamPause_CheckedChanged(object sender, EventArgs e)
        {
            if(m_Cfg != null )
            {
                m_bInCameraTestPause = materialCheckBoxInCamPause.Checked;
            }
        }

        private void materialCheckBoxOutCamPause_CheckedChanged(object sender, EventArgs e)
        {
            if( m_Cfg != null )
            {
                m_bOutCameraTestPause = materialCheckBoxOutCamPause.Checked;
            }
        }
    }

    public class CArea
    {
        public System.Drawing.Point start;
        public System.Drawing.Point end;

        public CArea()
        {
            start = new System.Drawing.Point(0, 0);
            end = new System.Drawing.Point( 0,0);
        }

        public override String ToString()
        {
            return start.ToString() + " - " + end.ToString();
        }
    }

    public class CV2Cfg {
        
        public Int32 m_nInCamera;
        public Int32 m_nInCameraPara;
        public Int32 m_nInCameraDelay;

        public Int32 m_nOutCamera;
        public Int32 m_nOutCameraPara;
        public Int32 m_nOutCameraDelay;

        public Int32 m_ModbusPort;
        public CArea m_AreaIn;  //// 进板识别区域
        public CArea m_AreaOut; //// 出板识别区域

        public Int32 m_iFeatureAreaLenth; //// 包含特征区域的延长长度，一个特征区域周围扩展多少像素，进行定位识别

        public double m_fCvFac;   //// 对比系数
        public List<CArea> m_CAreaListInTop;
        public List<String> m_CAreaMatInTop;

        public List<CArea> m_CAreaListInBtm;
        public List<String> m_CAreaMatInBtm;

        public List<CArea> m_CAreaListOutTop;
        public List<String> m_CAreaMatOutTop;

        public List<CArea> m_CAreaListOutBtm;
        public List<String> m_CAreaMatOutBtm;

        public List<CArea> m_CAreaListInTop2;
        public List<CArea> m_CAreaListInBtm2;
        public List<CArea> m_CAreaListOutTop2;
        public List<CArea> m_CAreaListOutBtm2;

        public Boolean m_bUseNewVisualTest;

        public CV2Cfg()
        {
            m_nInCamera = 0;
            m_nInCameraPara = 0;
            m_nInCameraDelay = 3;

            m_nOutCamera = 0;
            m_nOutCameraPara = 0;
            m_nOutCameraDelay = 3;

            m_iFeatureAreaLenth = 150;  //// 4650x3480 的图片，150像素大概是30mm左右。4000像素对应800mm，每毫米5个像素，150个像素大概30mm

            m_ModbusPort = 502;
            m_fCvFac = 0.8;
            m_AreaIn = new CArea();
            m_AreaOut = new CArea();

            m_CAreaListInTop = new List<CArea>();
            m_CAreaMatInTop = new List<String>();
            m_CAreaListInBtm = new List<CArea>();
            m_CAreaMatInBtm = new List<String>();

            m_CAreaListOutTop = new List<CArea>();
            m_CAreaMatOutTop = new List<String>();
            m_CAreaListOutBtm = new List<CArea>();
            m_CAreaMatOutBtm = new List<String>();

            m_CAreaListInTop2 = new List<CArea>();
            m_CAreaListInBtm2 = new List<CArea>();
            m_CAreaListOutTop2 = new List<CArea>();
            m_CAreaListOutBtm2 = new List<CArea>();

            m_bUseNewVisualTest = true;
        }

        public String fnPara2String()
        {
            String rtn = "";

            rtn ="进板摄像头 ：" + m_nInCamera.ToString() + ", " + m_nInCameraPara.ToString() + "\n" +
                "出板摄像头 ：" + m_nOutCamera.ToString() + ", " + m_nOutCameraPara.ToString() + "\n" +
                "进板区域：" + m_AreaIn.ToString() + "\n" +
                ", 出板区域：" + m_AreaOut.ToString() + "\n" +
                ", 进板正面特征：" + m_CAreaListInTop.Count.ToString() + "\n" +
                ", 进板反面特征" + m_CAreaListInBtm.Count.ToString() + "\n" +
            ", 出板正面特征：" + m_CAreaListOutTop.Count.ToString() + "\n" +
            ", 出板反面特征" + m_CAreaListOutBtm.Count.ToString() + "。";

            for(int ii = 0; ii < m_CAreaListInTop.Count; ii++ )
            {
                rtn += "\n  进板正面特征 " + ii.ToString() + " " + m_CAreaListInTop[ii].ToString();
            }
            for (int ii = 0; ii < m_CAreaListInBtm.Count; ii++)
            {
                rtn += "\n  进板反面特征 " + ii.ToString() + " " + m_CAreaListInBtm[ii].ToString();
            }
            for (int ii = 0; ii < m_CAreaListOutTop.Count; ii++)
            {
                rtn += "\n  出板正面特征 " + ii.ToString() + " " + m_CAreaListOutTop[ii].ToString();
            }
            for (int ii = 0; ii < m_CAreaListOutBtm.Count; ii++)
            {
                rtn += "\n  出板反面特征 " + ii.ToString() + " " + m_CAreaListOutBtm[ii].ToString();
            }

            return rtn;
        }
        
        public String fnPara2String2()
        {
            String rtn = "";

            rtn = m_CAreaListInTop.Count.ToString() +
                "," + m_CAreaListInBtm.Count.ToString() +
            "," + m_CAreaListOutTop.Count.ToString() +
            "," + m_CAreaListOutBtm.Count.ToString() +
            ", in:" + m_AreaIn.ToString() + ", out:" + m_AreaOut.ToString();
            return rtn;
        }
    
        public Boolean fnTestConfig()
        {
            if(m_nInCamera == m_nOutCamera )
            {
                MessageBox.Show("进板摄像头和出板摄像头，不可以选择同一个摄像头！");
                return false;
            }


            return true;
        }
    }



}
