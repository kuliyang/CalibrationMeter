using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;  //串口serialport类
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Collections;

namespace insForm
{
    class Kei2010
    {

        #region 变量
        public SerialPort comm = new SerialPort();
        public StringBuilder builder = new StringBuilder();//存储读取到的数据，避免在事件处理方法中反复的创建，定义到外面。   


        public bool Closing = false;//是否正在关闭串口，执行Application.DoEvents，并阻止再次invoke     
        public bool Listening = false; //监听串口读数
        public bool WriteEvent = false; //串口写事件标识
        //数字表用
        public int ReadStyle = 9;//读数类型。0、电压；1、预留；2、预留
        public double ratio = 1;//转换比率          

        #endregion

        #region 通用函数

        //构造函数
        public Kei2010()
        {
            ReadStyle = 9;//读数方式
            ratio = 1;//转换比例
        }


        /// <summary>
        /// 串口写数据
        /// </summary>
        /// <param name="str">要写的数据</param>
        /// <returns>是否发送成功</returns>
        public bool CommWrite(string str)
        {
            if (GlobalVariable.GDemoState)//演示
                return true;
            //判断命令
            if (str.Length == 0)
                return false;
            //保证串口打开
            if (!comm.IsOpen)
                return false;
            //监听状态则等待
            while (Listening)
                Application.DoEvents();
            //发送数据
            try
            {
                WriteEvent = true;//标识正在写
                comm.DiscardInBuffer();  //抛弃输入缓存
                comm.DiscardOutBuffer(); //抛弃输出缓存
                comm.WriteLine(str);
            }
            finally
            {
                WriteEvent = false;
            }
            return true;
        }

        /// <summary>
        /// 串口数据接收函数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void comm_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            if (GlobalVariable.GDemoState)//演示
                return;

            if (Closing) return;//如果正在关闭，忽略操作，直接返回，尽快的完成串口监听线程的一次循环 
            while (WriteEvent)//如果正在写
            {
                Application.DoEvents();
            }
            //
            try
            {
                Listening = true;//正在监听
                int n = comm.BytesToRead;//先记录下来，避免某种原因，人为的原因，操作几次之间时间长，缓存不一致  
                byte[] buf = new byte[n];//声明一个临时数组存储当前来的串口数据  
                comm.Read(buf, 0, n);//读取缓冲数据   读取接收到的数据
                //直接按ASCII规则转换成字符串  
                builder.Append(Encoding.ASCII.GetString(buf));//存到StringBuilder
            }
            finally
            {
                Listening = false;//关闭监听 
            }
        }

        /// <summary>
        /// 判断串口接收到的字符串是否为数值
        /// </summary>
        /// <param name="str">要判断的字符串</param>
        /// <returns></returns>
        public bool IsNumeric(string str)
        {
            char[] s = str.ToCharArray();
            int p = 0;
            for (int i = 0; i < str.Length; i++)
            {
                if (s[i] == '+' || s[i] == '-' || s[i] == 'E' || s[i] == ' ')
                {
                    continue;
                }
                if (s[i] == '.' && p == 0)
                {
                    p++;
                    continue;
                }
                if (s[i] < '0' || s[i] > '9')
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 延时函数,需要延时多少秒
        /// </summary>
        /// <param name="delayTime"></param>
        /// <returns></returns>
        public bool Delay(double delayTime)
        {
            long startTimer = Environment.TickCount;
            while (Math.Abs(Environment.TickCount - startTimer) < delayTime * 1000)
            {
                Application.DoEvents();
            }
            return true;
        }

        #endregion

        #region 打开关闭串口与初始化

        /// <summary>
        /// 打开串口
        /// </summary>
        /// <returns></returns>
        public bool OpenComm()
        {
            if (GlobalVariable.GDemoState)//演示
                return true;

            if (comm == null)
                return false;
            //
            if (!comm.IsOpen)//打开串口
            {
                try
                {
                    comm.Open();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "串口错误!");
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// 关闭串口
        /// </summary>
        public void CloseComm()
        {
            if (GlobalVariable.GDemoState)//演示
                return;
            //
            
                if (comm != null)
                {
                    if (comm.IsOpen)
                    {
                        Closing = true;//是否正在关闭串口，执行Application.DoEvents，并阻止再次invoke
                        while (Listening) Application.DoEvents();//处理队列中的信息
                        CommWrite(":syst:loc\r");//打开本地模式，关闭远程程控模式
                        comm.Close();

                        Closing = false;
                    }
                }
            }

        /// <summary>
        /// 释放串口
        /// </summary>
        public void Dispose()
        {
            if (GlobalVariable.GDemoState)//演示
                return;

            if (comm != null)//说明创建有实例则释放串口
                comm.Dispose();
        }


        //打开串口
        public bool OpenComm(string name)
        {
            return true;
        }
        //关闭串口
        public void CloseComm(string name)
        {
            return;
        }

        /// <summary>
        /// 初始化串口
        /// 串口名可以用cbx控件设置
        /// </summary>
        public void InitCom(string comPort)
        {
            if (GlobalVariable.GDemoState)//演示
                return;
            //
            if (comPort == "")
            {
                MessageBox.Show("串口不可为空", "错误");
                return;
            }
            if (!comm.IsOpen)//在串口关闭状态下设置串口
            {
                comm.BaudRate = 9600;
                comm.DataBits = 8;
                comm.PortName = comPort;
                comm.Parity = Parity.None;
                comm.StopBits = StopBits.One;
                comm.NewLine = "\r";
                comm.DataReceived += comm_DataReceived;//添加事件监听
                //comm.DataReceived += new SerialDataReceivedEventHandler(comm_DataReceived);
            }

        }




        /// <summary>
        /// 获取有效的COM口
        /// </summary>
        /// <returns></returns>
        public static string[] ActivePorts()
        {
            ArrayList activePorts = new ArrayList();
            foreach (string pname in SerialPort.GetPortNames())
            {
                activePorts.Add(Convert.ToInt32(pname.Substring(3)));
            }
            activePorts.Sort();
            string[] mystr = new string[activePorts.Count];
            int i = 0;
            foreach (int num in activePorts)
            {
                mystr[i++] = "COM" + num.ToString();
            }
            return mystr;
        }

        #endregion







        /// <summary>
        /// 读取数字标示值
        /// 读数类型：0:直流电压档；1:交流电压档；2:直流电流档；3：交流电流档；4：电阻档
        /// 需要判断读到的数知否有效
        /// </summary>
        /// <param name="i">读数类型：0:直流电压档；1:交流电压档；2:直流电流档；3：交流电流档；4：电阻档</param>
        /// <returns></returns>
        public double CommRead(int i)
        {
            if (GlobalVariable.GDemoState)//演示
                return GlobalVariable.GDemoValue;
            //
            string str = "", fh = "";
            for (int e = 0; e < 3; e++)//读取3次
            {
                CommWrite("read?\r"); //发送读取命令
                Delay(2);
                str = builder.ToString();
                builder.Length = 0;
                if (str.Length != 0)
                {
                    if (str.IndexOf("DCL") < 0)
                    {
                        if (str.IndexOf("\r") > 0)
                        {
                            break;
                        }
                    }
                }
            }
            //没有收到数据
            if (str.Length == 0)
                return -9999;
            //接收错误
            if (str.IndexOf("DCL") > 0)
                return -9998;
            //结束位判断
            int n = str.IndexOf("\r");
            if (n <= 0)
                return -9997;
            //将str分为符号位与数值
            fh = str.Substring(0, 1);//获得符号位
            str = str.Substring(1, n - 1);//获得数字表数值
            if (IsNumeric(str))
            {
                double value;
                try
                {
                    if (i == 0)
                        value = double.Parse(str) * ratio;  //kei2010表读取的是V，转换为mv
                    else
                        value = double.Parse(str);
                }
                catch
                {
                    return -9995;//数值转换错误
                }
                if (fh == "-")
                    value = value * -1;

                return value;
            }
            else
            {
                return -9996;//含有非法字符
            }
        }

        /// <summary>
        /// Kei2010清除所有的位和注册事件,仪表初始化
        /// </summary>
        public void Kei2010_Init()
        {
            //*CLS   清除所有的位和注册事件
            //:INIT   进入互触发状态
            //:SYST:REM 将仪表处于RS-232或者以太网远程控制模式；除本地键以外，前面板按键禁用
            CommWrite("*RST\r;*CLS\r;:INIT\r;:SYST:REM\r");
            // CommWrite("func \"volt:dc\"\r");
        }

        /// <summary>
        /// 仪表初始化设置
        /// 编程手册37页
        /// 接下来要读书CommRead
        /// DCV,ACV,电阻档都在一起，不需要改线
        /// DCI,ACI在一起
        /// 四线电阻一个口
        /// </summary>
        /// <param name="i">0:直流电压档；1:交流电压档；2:直流电流档；3：交流电流档；4：电阻档</param>
        /// <param name="s">设置读数速度:快0.2/中10/慢100</param>
        public void Init(int i, string s, int filter)
        {

            if (GlobalVariable.GDemoState)//演示
                return;
            //

            if (ReadStyle == i)//如果等于9返回
                return;
            //
            ReadStyle = i;//读数仪表 类型   
            if (ReadStyle == 0)//0:直流电压档 最大1000V peak
            {
                CommWrite("*RST\r;*CLS\r;:INIT\r;:SYST:REM\r");//rem远程模式
                CommWrite("func \'volt:dc\'\r");  //电压档
                //CommWrite("func\'Volt:DC:Rang:AUTO\'\r");
                //设置滤波次数
                CommWrite("VOLT:DC:AVER:STAT 1\r");//开启滤波
                CommWrite("VOLT:DC:AVER:TCON REP\r");//滤波类型
                CommWrite("VOLT:DC:AVER:COUN " + filter.ToString() + "\r");//滤波值（1-100）1快5中10慢
                //CommWrite("VOLT:DC:AVER:COUN " + GlobalVariable.Gfilter + "\r");//滤波值（1-100）1快5中10慢
                //设置读数速度
                CommWrite("VOLT:DC:NPLC " + s + "\r");//设置读数速度（0.01-10）
            }
            else if (ReadStyle == 1) //1:交流电压档 最大范围750V rms, 1000V peak, 8×107V•Hz
            {
                CommWrite("*RST\r;*CLS\r;:INIT\r;:SYST:REM\r");//rem远程模式
                CommWrite("func \'volt:ac\'\r");  //交流电压档volt：路径配置到电压单元
                //CommWrite("func\'Volt:AC:Rang:AUTO\'\r");
                //设置滤波次数
                CommWrite("Volt:AC:AVER:STAT 1\r");//开启滤波
                CommWrite("Volt:AC:AVER:TCON REP\r");//滤波类型
                CommWrite("Volt:AC:AVER:COUN " + filter.ToString() + "\r");//滤波值（1-100）1快5中10慢
                //CommWrite("Volt:AC:AVER:COUN " + GlobalVariable.Gfilter + "\r");//滤波值（1-100）1快5中10慢
                //设置读数速度
                CommWrite("Volt:AC:NPLC " + s + "\r");//设置读数速度（0.01-10）
            }
            else if (ReadStyle == 2)//2:直流电流档 最大范围：3A dc, 250V
            {
                CommWrite("*RST\r;*CLS\r;:INIT\r;:SYST:REM\r");//rem远程模式
                CommWrite("func \'CURR:DC\'\r");  //电流档curr：路径配置到电流单元
                //CommWrite("func \'CURR:DC:rang:auto\'\r");
                //设置滤波次数
                CommWrite("CURR:DC:AVER:STAT 1\r");//开启滤波
                CommWrite("CURR:DC:AVER:TCON REP\r");//滤波类型
                CommWrite("CURR:DC:AVER:COUN " + filter.ToString() + "\r");//滤波值（1-100）1快5中10慢
                //CommWrite("CURR:DC:AVER:COUN " + GlobalVariable.Gfilter + "\r");//滤波值（1-100）1快5中10慢
                //设置读数速度
                CommWrite("CURR:DC:NPLC " + s + "\r");//设置读数速度（0.01-10）
            }
            else if (ReadStyle == 3)//3：交流电流档 最大范围 3A rms, 250V
            {
                CommWrite("*RST\r;*CLS\r;:INIT\r;:SYST:REM\r");//rem远程模式
                CommWrite("func \'curr:ac\'\r");
                //CommWrite("func \'curr:ac:rang:auto\'\r");//设置量程为自动
                //设置滤波次数
                CommWrite("curr:ac:AVER:STAT 1\r");//使能滤波
                CommWrite("curr:ac:AVER:TCON REP\r");//滤波类型:repeat
                CommWrite("curr:ac:AVER:COUN " + filter.ToString() + "\r");//滤波值（1-100）1快5中10慢
                //CommWrite("curr:ac:AVER:COUN " + GlobalVariable.Gfilter + "\r");//滤波值（1-100）1快5中10慢
                //设置读数速度
                CommWrite("curr:ac:NPLC " + s + "\r");//设置读数速度（0.01-10）
            }
            else if (ReadStyle == 4)//4：电阻档 最大12MΩ
            {
                CommWrite("func \'res\'\r");  //二线电阻档
                //设置滤波次数
                CommWrite("RES:AVER:STAT 1\r");
                CommWrite("RES:AVER:TCON REP\r");//滤波类型
                CommWrite("RES:AVER:COUN " + filter.ToString() + "\r");//滤波值（1-10）
                //CommWrite("RES:AVER:COUN " + GlobalVariable.Gfilter + "\r");//滤波值（1-10）
                //设置读数速度
                CommWrite("RES:NPLC " + s + "\r");
            }

            Delay(5);
        }

        /// <summary>
        /// 本地状态(RS-232 only)
        /// 关闭通讯状态
        /// </summary>
        public void Uinit()
        {
            if (GlobalVariable.GDemoState)//演示
                return;
            CommWrite(":syst:loc\r");
        }

        /// <summary>
        /// 远程状态
        /// 启动远程通讯
        /// </summary>
        public void Rem()
        {
            if (GlobalVariable.GDemoState)//演示
                return;
            //*CLS   清除所有的位和注册事件
            //:INIT   进入互触发状态
            //:SYST:REM 将仪表处于RS-232或者以太网远程控制模式；除本地键以外，前面板按键禁用
            CommWrite("*RST\r;*CLS\r;:INIT\r;:SYST:REM\r");
        }

        //换挡(func：档)
        public void SetFunc(int func)
        {

        }
        //写参数命令（成功返回true）
        public bool CommWritePara(string str)
        {
            return true;
        }

        /// <summary>
        /// 数字表通讯测试
        /// </summary>
        /// <param name="testData"></param>
        /// <returns></returns>
        public bool Test(ref double testData)
        {
            if (GlobalVariable.GDemoState)//演示
            {
                testData = GlobalVariable.GDemoValue;
                return true;
            }
            //
            double shu = 0; //用于存储读取电压数值
            for (int i = 0; i < 3; i++)
            {
                shu = CommRead(0); //读取电压数值
                if (shu > -9000)
                    break;
            }
            testData = shu;
            if (shu > -9000)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 测试电压档
        /// </summary>
        /// <returns></returns>
        public bool MeterTryVoltage(out string s)
        {
            double mes;

            CommWrite("*RST\r;*CLS\r;:INIT\r;:SYST:REM\r");
            Delay(2);

            CommWrite("VOLT:DC:FILT 1\r");//打开滤波器
            CommWrite("func \"volt:dc\"\r");
            Delay(2);
            CommWrite("VOLT:DC:NPLC 0.2\r");//电力线周期数；将直流电压积分时间设为周期的0.2倍

            CommWrite("read?\r");
            mes = CommRead(0);
            if (mes == -9999)//数值处理方法
            {
                MessageBox.Show("没有接收到数据，请检查：\r\n① 设备是否上电；\r\n② 数字多用表型号是否正确；\r\n③ 端口号是否正确；\r\n" +
                    "④ 数字多用表串口是否打开；\r\n⑤ 通讯线是否连接正确；\r\n⑥ 通讯线是否损坏。");
                s = "";
                return false;
            }
            else if (mes == -9998)
            {
                MessageBox.Show("接收错误");
                s = "";
                return false;
            }
            else if (mes == -9997)
            {
                MessageBox.Show("无结束位");
                s = "";
                return false;
            }
            else if (mes == -9996)
            {
                MessageBox.Show("含有非法字符");
                s = "";
                return false;
            }

            CommWrite("syst:loc\r");
            s = "通讯测试成功";
            return true;
        }



        #region 数字表、转换开关一体化专用函数
        //开关换向
        public int DTong(bool b)
        {
            return 1;
        }
        //开关换线
        public int Reverse(bool b)
        {
            return 1;
        }
        //开关走位
        public int Go(int i)
        {
            return 1;
        }
        //开关测试
        public int Test()
        {
            return 1;
        }

        #endregion





    }
}
