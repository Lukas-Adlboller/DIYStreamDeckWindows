using System.Diagnostics;
using System.IO.Ports;
using System.Management;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using NAudio;
using NAudio.CoreAudioApi;

namespace ShortcutButtonApp
{
  public partial class MainForm : Form
  {
    public MMDevice MainMic;

    private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
    private const int WM_APPCOMMAND = 0x319;

    public const int KEYEVENTF_EXTENTEDKEY = 1;
    public const int KEYEVENTF_KEYUP = 0;
    public const int VK_MEDIA_NEXT_TRACK = 0xB0;// code to jump to next track
    public const int VK_MEDIA_PLAY_PAUSE = 0xB3;// code to play or pause a song
    public const int VK_MEDIA_PREV_TRACK = 0xB1;// code to jump to prev track

    [DllImport("user32.dll")]
    public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, IntPtr extraInfo);

    private void Mute()
    {
      SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
          (IntPtr)APPCOMMAND_VOLUME_MUTE);
    }

    private void Goodbye()
    {
      Process.Start("shutdown", "/s /t 0");
    }


    public static string dataBuffer = string.Empty;

    struct ComPort // custom struct with our desired values
    {
      public string name;
      public string vid;
      public string pid;
      public string description;
    }

    public static SerialPort serialPort;
    public static bool isInitialized = false;

    public enum DeviceCommands { MuteOrUnmuteTS, MuteOrUnmuteDC, PlayValorant, PCShutdown, ClipCapture, SilentMode, GlobalMute, PlayPause, SkipSong, NoCmd };

    private const string vidPattern = @"VID_([0-9A-F]{4})";
    private const string pidPattern = @"PID_([0-9A-F]{4})";

    public DeviceCommands GetCommand(string data)
    {
      if (data.IndexOf("mtntmts") != -1)
      {
        return DeviceCommands.MuteOrUnmuteTS;
      }
      else if (data.IndexOf("mtntmdc") != -1)
      {
        return DeviceCommands.MuteOrUnmuteDC;
      }
      else if (data.IndexOf("plyvlrnt") != -1)
      {
        return DeviceCommands.PlayValorant;
      }
      else if (data.IndexOf("pcshtdwn") != -1)
      {
        return DeviceCommands.PCShutdown;
      }
      else if (data.IndexOf("clpcptr") != -1)
      {
        return DeviceCommands.ClipCapture;
      }
      else if (data.IndexOf("slntmd") != -1)
      {
        return DeviceCommands.SilentMode;
      }
      else if(data.IndexOf("glblmt") != -1)
      {
        return DeviceCommands.GlobalMute;
      }
      else if (data.IndexOf("skpsng") != -1)
      {
        return DeviceCommands.SkipSong;
      }
      else if (data.IndexOf("plyps") != -1)
      {
        return DeviceCommands.PlayPause;
      }
      else
      {
        return DeviceCommands.NoCmd;
      }
    }

    private static List<ComPort> GetSerialPorts()
    {
      using (var searcher = new ManagementObjectSearcher
          ("SELECT * FROM WIN32_SerialPort"))
      {
        var ports = searcher.Get().Cast<ManagementBaseObject>().ToList();
        return ports.Select(p =>
        {
          ComPort c = new ComPort();
          c.name = p.GetPropertyValue("DeviceID").ToString();
          c.vid = p.GetPropertyValue("PNPDeviceID").ToString();
          c.description = p.GetPropertyValue("Caption").ToString();

          Match mVID = Regex.Match(c.vid, vidPattern, RegexOptions.IgnoreCase);
          Match mPID = Regex.Match(c.vid, pidPattern, RegexOptions.IgnoreCase);

          if (mVID.Success)
            c.vid = mVID.Groups[1].Value;
          if (mPID.Success)
            c.pid = mPID.Groups[1].Value;

          return c;

        }).ToList();
      }
    }

    private string GetPortName()
    {
      List<ComPort> ports = GetSerialPorts();
      ComPort com = ports.FindLast(c => c.vid.Equals("2341") && c.pid.Equals("0001"));

      return com.name;
    }

    private void onSerialDataRecieve(object sender, SerialDataReceivedEventArgs args)
    {
      string strData = serialPort.ReadExisting();

      dataBuffer += strData;

      if (!dataBuffer.EndsWith("\r\n"))
      {
        return;
      }

      if (dataBuffer.IndexOf("dvcrdy") == -1 && isInitialized == false)
      {
        isInitialized = true;
      }

      if (isInitialized != true || string.IsNullOrEmpty(dataBuffer))
      {
        return;
      }

      DeviceCommands deviceCommand = GetCommand(dataBuffer);
      switch (deviceCommand)
      {
        case DeviceCommands.MuteOrUnmuteTS:
          SendKeys.SendWait("{SUBTRACT}");
          break;
        case DeviceCommands.MuteOrUnmuteDC:
          SendKeys.SendWait("{ADD}");
          break;
        case DeviceCommands.PlayValorant:
          ProcessStartInfo info = new ProcessStartInfo("D:\\Games\\Riot Games\\Riot Client\\RiotClientServices.exe");
          info.Arguments = " --launch-product=valorant --launch-patchline=live";
          info.WorkingDirectory = "D:\\Games\\Riot Games\\Riot Client\\";
          info.UseShellExecute = true;
          info.Verb = "runas";
          Process.Start(info);
          break;
        case DeviceCommands.PCShutdown:
          this.BeginInvoke(() => { Goodbye(); });
          break;
        case DeviceCommands.ClipCapture:
          SendKeys.SendWait("%{F10}");
          break;
        case DeviceCommands.SilentMode:
          this.BeginInvoke(() => { Mute(); });
          break;
        case DeviceCommands.GlobalMute:
          this.BeginInvoke(() => { MainMic.AudioEndpointVolume.Mute = !MainMic.AudioEndpointVolume.Mute; });
          break;
        case DeviceCommands.PlayPause:
          this.BeginInvoke(() => { keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero); });
          break;
        case DeviceCommands.SkipSong:
          this.BeginInvoke(() => { keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero); });
          break;
        case DeviceCommands.NoCmd:
          break;
      }

      dataBuffer = string.Empty;
    }

    public MainForm()
    {
      InitializeComponent();
    }

    private void MainForm_Load(object sender, EventArgs e)
    {
      notifyIcon.Visible = true;
      notifyIcon.ShowBalloonTip(10);
      string portName = GetPortName();

      while ((portName = GetPortName()) == null)
      {
      }

      serialPort = new SerialPort(portName, 9600);
      serialPort.DataReceived += new SerialDataReceivedEventHandler(onSerialDataRecieve);

      var enumerator = new MMDeviceEnumerator();
      MainMic = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);

      try
      {
        serialPort.Open();

        if (!serialPort.IsOpen)
        {
          Environment.Exit(-1);
        }
      }
      catch (IOException ex)
      {
        Environment.Exit(-1);
      }
    }
  }
}