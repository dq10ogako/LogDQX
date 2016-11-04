//
// LogDQX.cs
//

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Drawing;
using System.Threading;
using System.ComponentModel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Timers;

using FNF.Utility;
using FNF.Controls;
using FNF.XmlSerializerSetting;
using FNF.BouyomiChanApp;

public static class Lo {
  [DllImport("LogDQX.dll", EntryPoint="LoInitialize")]
  public static extern int Initialize(int pid);
  [DllImport("LogDQX.dll", EntryPoint="LoCheck")]
  public static extern int Check();
  [DllImport("LogDQX.dll", EntryPoint="LoRead")]
  public static extern int Read(IntPtr p);
  [DllImport("LogDQX.dll", EntryPoint="LoReadW")]
  public static extern int ReadW(IntPtr p);
  [DllImport("LogDQX.dll", EntryPoint="LoProcess")]
  public static extern int Process();
}

namespace Plugin_LogDQX {
  public struct Voice {
    public int Speed, Volume, Type;
  }

  public class Plugin_LogDQX : IPlugin {
    public string Name { get { return "DQX チャット読み上げ"; } }
    public string Version { get { return "2016/10/23版"; } }
    public string Caption { get { return "DQX チャット読み上げ"; } } 
    public ISettingFormData SettingFormData {
      get { return _SettingFormData; }
    }

    const int CHAT_MAX = 2;

    private Settings_LogDQX _Settings;
    private SettingFormData_LogDQX _SettingFormData;
    private string _SettingFile =
      Base.CallAsmPath + Base.CallAsmName + ".setting";

    private ToolStripButton _Button;
    private ToolStripSeparator _Separator;

    private bool _Enable = true;
    private Bitmap _IconEnable;
    private Bitmap _IconDisable;

    internal void SetNextAlart() {
      for (int i = 0; i < CHAT_MAX; i++) {
        if ((_Settings.Chat[i].Speed < 50) || (_Settings.Chat[i].Speed > 500)) {
          _Settings.Chat[i].Speed = -1;
        }
        if ((_Settings.Chat[i].Volume < 0) || (_Settings.Chat[i].Volume > 100)) {
          _Settings.Chat[i].Volume = -1;
        }
        if (_Settings.Chat[i].Type < 0) {
          _Settings.Chat[i].Type = 0;
        }
      }
    }

    public Plugin_LogDQX() {
      _IconEnable = Properties.Log.ImgLog;
      _IconDisable = BitmapGrayScale.Convert(_IconEnable);
    }

    public void Begin() {
      _Settings = new Settings_LogDQX(this);
      _Settings.Load(_SettingFile);
      _SettingFormData = new SettingFormData_LogDQX(_Settings);

      _Separator = new ToolStripSeparator();
      Pub.ToolStrip.Items.Add(_Separator);
      _Button = new ToolStripButton(Properties.Log.ImgLog);
      _Button.ToolTipText = "DQX チャット読み上げ";
      _Button.Click += ButtonClick;
      Pub.ToolStrip.Items.Add(_Button);

      _Enable = true;
      StartWatch();
    }

    public void End() {
      _Settings.Save(_SettingFile);

      if (_Separator != null) {
        Pub.ToolStrip.Items.Remove(_Separator);
        _Separator.Dispose();
        _Separator = null;
      }
      if (_Button != null) {
        Pub.ToolStrip.Items.Remove(_Button);
        _Button.Dispose();
        _Button = null;
      }

      _Enable = false;
      StopWatch();
    }

    private void ButtonClick(object sender, EventArgs e) {
      _Enable = !_Enable;
      if (_Enable) {
        _Button.Image = _IconEnable;
        StartWatch();
      } else {
        _Button.Image = _IconDisable;
        StopWatch();
      }
    }

    private bool _First = true;
    private int _Pid = 0;
    private System.Timers.Timer _Timer = null;

    private void StartWatch() {
      StopWatch();

      if (_Timer == null) {
        _Timer = new System.Timers.Timer();
        _Timer.Enabled = true;
        _Timer.AutoReset = true;
        _Timer.Interval = 1000;
        _Timer.Elapsed += new ElapsedEventHandler(OnTimer);
      }
      _Timer.Start();
    }

    private void StopWatch() {
      if (_Timer != null) {
        _Timer.Stop();
      }
      _First = true;
      _Pid = 0;
    }

    private void OnTimer(object source, ElapsedEventArgs e) {
      int r;

      if (_Pid == 0) {
        _Pid = Lo.Process();

        if (_Pid == 0) {
          return;
        }

        r = Lo.Initialize(_Pid);
        if (r == 0) {
          _Pid = 0;
          return;
        }
      }

      r = Lo.Check();
      if (r == 0) {
        _Pid = 0;
        return;
      }

      ReadLog();
    }

    private IntPtr p = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(byte)) * 524288);

    private void ReadLog() {
      if (_Pid == 0) {
        return;
      }

      int c = Lo.ReadW(p);
      if (c < 2) {
        return;
      }

      if (_First) {
        _First = false;
        return;
      }

      string s = Marshal.PtrToStringUni(p);
      string [] lines;
      lines = s.Split('\n');
      foreach (string line in lines) {
        string [] t;
        t = line.Split('\t');
        if (t.Length > 7) {
          string name = t[5];
          string chat = t[7];
          if (name != "") {
            Talk(name, chat);
          }
        }
      }
    }

    private void Talk(string n, string l) {
      int speed = -1;
      int volume = -1;
      int type = 0;

      if (_Settings.Name) {
        speed = _Settings.Chat[1].Speed;
        volume = _Settings.Chat[1].Volume;
        type = _Settings.Chat[1].Type;
        Pub.AddTalkTask(n, speed, volume, (FNF.Utility.VoiceType)type);
      }

      speed = _Settings.Chat[0].Speed;
      volume = _Settings.Chat[0].Volume;
      type = _Settings.Chat[0].Type;
      Pub.AddTalkTask(l, speed, volume, (FNF.Utility.VoiceType)type);
    }


    public class Settings_LogDQX : SettingsBase {
      public Voice [] Chat = new Voice[CHAT_MAX];
      public bool Name = true;

      internal Plugin_LogDQX Plugin;
      public Settings_LogDQX() {
        _Initialize();
      }

      public Settings_LogDQX(Plugin_LogDQX p) {
        Plugin = p;
        _Initialize();
      }
      public override void ReadSettings() {
      }
      public override void WriteSettings() {
        Plugin.SetNextAlart();
      }

      private void _Initialize() {
        for (int i = 0; i < CHAT_MAX; i++) {
          Chat[i].Speed = -1;
          Chat[i].Volume = -1;
          Chat[i].Type = 0;
        }
        Name = true;
      }
    }


    public class SettingFormData_LogDQX : ISettingFormData {
      Settings_LogDQX _Setting;

      public string Title { get { return _Setting.Plugin.Name; } }
      public bool ExpandAll { get { return false; } }
      public SettingsBase Setting { get { return _Setting; } }

      public SettingFormData_LogDQX(Settings_LogDQX setting) {
        _Setting = setting;
        PBase = new SBase(_Setting);
      }

      public SBase PBase;
      public class SBase : ISettingPropertyGrid {
        Settings_LogDQX _Setting;
        public SBase(Settings_LogDQX setting) { _Setting = setting; }
        public string GetName() { return "設定"; }

        [Category ("本文設定")]
        [DisplayName("1) 速度")]
        [Description("読み上げの速度を 50 から 200 で指定してください。-1 を指定するとデフォルトの設定を使います。")]
        public int Speed0 { get { return _Setting.Chat[0].Speed; } set { _Setting.Chat[0].Speed = value; } }

        [Category ("本文設定")]
        [DisplayName("2) 音量")]
        [Description("読み上げの音量を 0 から 100 の数字で指定してください。-1 を指定するとデフォルトの設定を使います。")]
        public int Volume0 { get { return _Setting.Chat[0].Volume; } set { _Setting.Chat[0].Volume = value; } }

        [Category ("本文設定")]
        [DisplayName("3) 声質")]
        [Description("声のタイプを 1 以上の数値で指定してください。0 を指定するとデフォルトの設定を使います。")]
        public int Type0 { get { return _Setting.Chat[0].Type; } set { _Setting.Chat[0].Type = value; } }

        [Category ("名前設定")]
        [DisplayName("4) 名前を読み上げる")]
        [Description("発言者の名前を読み上げます。")]
        public bool Name { get { return _Setting.Name; } set { _Setting.Name = value; } }

        [Category ("名前設定")]
        [DisplayName("5) 速度")]
        [Description("読み上げの速度を 50 から 200 で指定してください。-1 を指定するとデフォルトの設定を使います。")]
        public int Speed1 { get { return _Setting.Chat[1].Speed; } set { _Setting.Chat[1].Speed = value; } }

        [Category ("名前設定")]
        [DisplayName("6) 音量")]
        [Description("読み上げの音量を 0 から 100 の数字で指定してください。-1 を指定するとデフォルトの設定を使います。")]
        public int Volume1 { get { return _Setting.Chat[1].Volume; } set { _Setting.Chat[1].Volume = value; } }

        [Category ("名前設定")]
        [DisplayName("7) 声質")]
        [Description("声のタイプを 1 以上の数値で指定してください。0 を指定するとデフォルトの設定を使います。")]
        public int Type1 { get { return _Setting.Chat[1].Type; } set { _Setting.Chat[1].Type = value; } }

      }
    }

  }
}
