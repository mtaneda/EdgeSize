// EdgeSize: Microsoft Edge をお気に入りのウィンドウサイズで起動する
// Main Form
//
// Copyright (C) 2023 by TANEDA M.
// All Rights Reserved.
// This code was designed and coded by TANEDA M.
//
// Last modified: 2023/03/29

using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Automation;
using System.Windows.Forms;

/// <summary>
///  メインフォーム
/// </summary>
public partial class MainForm : Form {
#region 定数
    /// <summary>
    ///  INIファイル名
    /// </summary>
    public static readonly string INI_FILE_NAME = "EdgeSize.ini";

    /// <summary>
    ///  デフォルト値：プロセス起動後の待ち時間
    /// </summary>
    public static readonly int DEFAULT_SLEEP_MS = 200;

    /// <summary>
    ///  デフォルト値：起動する実行ファイル
    /// </summary>
    public static readonly string DEFAULT_PATH = @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe";

    /// <summary>
    ///  デフォルト値：実行ファイルへの引数
    /// </summary>
    public static readonly string DEFAULT_ARG  = "--app=http://www.yahoo.co.jp/";

    /// <summary>
    ///  デフォルト値：起動したプロセスの名前
    /// </summary>
    public static readonly string DEFAULT_PROCESS_NAME = "msedge";

    /// <summary>
    ///  デフォルト値：対象となるウィンドウのクラス名
    /// </summary>
    public static readonly string DEFAULT_CLASS_NAME = "Chrome_WidgetWin_1";

    /// <summary>
    ///  デフォルト値：ウィンドウの識別子とするProperty(自国語)
    /// </summary>
    public static readonly string DEFAULT_NAME_PROPERTY_LOCAL = "アドレスと検索バー";

    /// <summary>
    ///  デフォルト値：ウィンドウの識別子とするProperty(英語)
    /// </summary>
    public static readonly string DEFAULT_NAME_PROPERTY = "address and search bar";

    /// <summary>
    ///  デフォルト値：Propertyの正規表現
    /// </summary>
    public static readonly string DEFAULT_REGEX = "[\"http://\"|\"https://\"]www\\.yahoo\\.co\\.jp.*";

    /// <summary>
    ///  デフォルト値：ウィンドウ位置X
    /// </summary>
    public static readonly int DEFAULT_X = 0;

    /// <summary>
    ///  デフォルト値：ウィンドウ位置Y
    /// </summary>
    public static readonly int DEFAULT_Y = 0;

    /// <summary>
    ///  デフォルト値：ウィンドウ幅
    /// </summary>
    public static readonly int DEFAULT_WIDTH = 960;

    /// <summary>
    ///  デフォルト値：ウィンドウ高
    /// </summary>
    public static readonly int DEFAULT_HEIGHT = 1024;

#endregion

#region Win32API
    public const int  HWND_NOTOPMOST =     -2;
    public const int  SWP_SHOWWINDOW = 0x0040;

    [DllImport("kernel32.dll")]
    public static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, uint nSize, string lpFileName);

    [DllImport("kernel32.dll")]
    public static extern int GetPrivateProfileInt(string lpAppName, string lpKeyName, int nDefault, string lpFileName);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int IsWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int cx, int cy, int uFlags);
#endregion

    /// <summary>
    ///  メインフォームのコンストラクタ
    /// </summary>
    public MainForm() {
        InitializeComponent();
    }

    /// <summary>
    ///  アドレスバーの内容からウィンドウハンドルを得る
    /// </summary>
    /// <param name="processName"></param>
    /// <param name="className"></param>
    /// <param name="namePropertyLocal"></param>
    /// <param name="nameProperty"></param>
    /// <param name="addressRegex"></param>
    /// <returns></returns>
   private IntPtr getHwndByAddress(string processName,
                                   string className,
                                   string namePropertyLocal,
                                   string nameProperty,
                                   string addressRegex) {
        //
        // 処理すべきプロセスを列挙する
        Process[] processes = Process.GetProcessesByName(processName);
        foreach (Process process in processes) {
            //
            // ウィンドウがないプロセスは除外（特に Edge は、バックグラウンドプロセスをたくさん起動するため）
            if(process.MainWindowHandle == IntPtr.Zero)
                continue;
            AutomationElementCollection roots = AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children,                                                 // スコープ
                                                                                      new AndCondition(                                                                       // 条件
                                                                                                       new PropertyCondition(AutomationElement.ProcessIdProperty, process.Id),//  プロセスID
                                                                                                       new PropertyCondition(AutomationElement.ClassNameProperty, className)  //  クラス名
                                                                                      ));
            foreach (AutomationElement rootElement in roots) {
                AutomationElement addressBar = rootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, namePropertyLocal));
                if(addressBar == null) {
                    addressBar = rootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, nameProperty));
                    if(addressBar == null)
                        continue;
                }

                ValuePattern v = (ValuePattern)addressBar.GetCurrentPattern(ValuePattern.Pattern);
                if(v.Current.Value != null && v.Current.Value != "")
                    if(Regex.IsMatch(v.Current.Value, addressRegex))
                        return (IntPtr)rootElement.Current.NativeWindowHandle;
            }
        }
        return IntPtr.Zero;
    }
    /// <summary>
    ///  ウィンドウサイズと位置を変更する
    /// </summary>
    /// <param name="hWnd"></param>
    /// <param name="point"></param>
    /// <param name="size"></param>
    /// <returns></returns>
    private bool setWindowPos(IntPtr hWnd, Point point, Size size) {
        if(IsWindow(hWnd) != 0)
            return SetWindowPos(hWnd, HWND_NOTOPMOST, point.X, point.Y, size.Width, size.Height, SWP_SHOWWINDOW);
        else
            return false;
    }

#region イベントハンドラ
    /// <summary>
    ///  メインフォーム構築時のハンドラ
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Form1_Load(object sender, EventArgs e) {
        Process proc;
        StringBuilder sb;
        IntPtr hWnd;
        string iniPath, sect, path, arg, procName, className, namePropertyLocal, nameProperty, addressRegex;
        int count, x, y, width, height, sleepMs;

        //
        // SystemセクションのCount分実行ファイルを起動する
        iniPath = Directory.GetParent(Assembly.GetExecutingAssembly().Location).ToString() + "\\" + INI_FILE_NAME;
        count = GetPrivateProfileInt("System", "Count", 2, iniPath);
        for(int i = 1; i <= count; i++) {
            sect = String.Format("App{0}", i);
            proc = new Process();

            sb = new StringBuilder(256);
            GetPrivateProfileString(sect, "Path", DEFAULT_PATH, sb, Convert.ToUInt32(sb.Capacity), iniPath);
            path = sb.ToString();
            GetPrivateProfileString(sect, "Arg", DEFAULT_ARG, sb, Convert.ToUInt32(sb.Capacity), iniPath);
            arg = sb.ToString();
            GetPrivateProfileString(sect, "ProcName", DEFAULT_PROCESS_NAME, sb, Convert.ToUInt32(sb.Capacity), iniPath);
            procName = sb.ToString();
            GetPrivateProfileString(sect, "ClassName", DEFAULT_CLASS_NAME, sb, Convert.ToUInt32(sb.Capacity), iniPath);
            className = sb.ToString();
            GetPrivateProfileString(sect, "NamePropertyLocal", DEFAULT_NAME_PROPERTY_LOCAL, sb, Convert.ToUInt32(sb.Capacity), iniPath);
            namePropertyLocal = sb.ToString();
            GetPrivateProfileString(sect, "NameProperty", DEFAULT_NAME_PROPERTY, sb, Convert.ToUInt32(sb.Capacity), iniPath);
            nameProperty = sb.ToString();
            GetPrivateProfileString(sect, "Address", DEFAULT_REGEX, sb, Convert.ToUInt32(sb.Capacity), iniPath);
            addressRegex = sb.ToString();

            proc.StartInfo.FileName = path;
            proc.StartInfo.Arguments = arg;
            proc.Start();
            proc.WaitForInputIdle();
            // アドレスバーの内容が取得できないことがあるので少し待つ
            // ※手抜きなのでメッセージループがこの時間ブロックします
            sleepMs = GetPrivateProfileInt("System", "Sleep", DEFAULT_SLEEP_MS, iniPath);
            System.Threading.Thread.Sleep(sleepMs);

            //
            // ウィンドウハンドルを取得します
            hWnd = getHwndByAddress(procName, className, namePropertyLocal, nameProperty, addressRegex);

            //
            // ウィンドウサイズと位置を変更します
            x = GetPrivateProfileInt(sect, "X", DEFAULT_X, iniPath);
            y = GetPrivateProfileInt(sect, "Y", DEFAULT_Y, iniPath);
            width = GetPrivateProfileInt(sect, "Width", DEFAULT_WIDTH, iniPath);
            height = GetPrivateProfileInt(sect, "Height", DEFAULT_HEIGHT, iniPath);
            if(!setWindowPos(hWnd, new Point(x, y), new Size(width, height)))
                MessageBox.Show("ウィンドウサイズ変更失敗", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        System.Environment.Exit(0);
    }
#endregion
}
