using System.Diagnostics;
using System.IO.Ports;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Google.Apis.YouTube.v3;
using NAudio.CoreAudioApi;
using SpotifyAPI.Web;
using SpotifyAPI.Web.Auth;
using Newtonsoft.Json;

namespace ShortcutButtonApp
{
  public class ApiCredentials
  {
    public string ApiName { get; set; }
    public string AppId { get; set; }
    public string AppSecret { get; set; }
    public string Token { get; set; }
  }

  public class JsonRoot
  {
    public IList<ApiCredentials> ApiCredentials { get; set; }
  }

  public partial class MainForm : Form
  {
    // DLL Function Imports
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, IntPtr extraInfo);

    // Keyboard Key Constants
    public const int KEYEVENTF_KEYDOWN = 0x00;
    public const int KEYEVENTF_EXTENTEDKEY = 0x01;
    public const int VK_MEDIA_NEXT_TRACK = 0xB0; // code to jump to next track
    public const int VK_MEDIA_PLAY_PAUSE = 0xB3; // code to play or pause a song

    // Microphone Handle and Constants
    public static MMDevice MainMic;
    private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
    private const int WM_APPCOMMAND = 0x319;

    // Main Sound Device
    private static MMDevice MainSound;

    // COM Device Regex Patterns
    private const string vidPattern = @"VID_([0-9A-F]{4})";
    private const string pidPattern = @"PID_([0-9A-F]{4})";

    // Google Stuff
    public static ApiCredentials ytApiCredentials;
    public static YouTubeService ytService;

    // Spotify Stuff
    // Get Auth Token: https://accounts.spotify.com/authorize?response_type=code&client_id=$CLIENT_ID&scope=$SCOPE&redirect_uri=$REDIRECT_URI
    // Get Refresh Token: curl -d client_id=$CLIENT_ID -d client_secret=$CLIENT_SECRET -d grant_type=authorization_code -d code=$AuthCode -d redirect_uri=http://localhost:5000/callback https://accounts.spotify.com/api/token
    private static EmbedIOAuthServer spotifyAuthServer;
    public static ApiCredentials spotifyApiCredentials;
    public static AuthorizationCodeRefreshResponse spotifyAuthResponse;
    public static SpotifyClient spotifyApi;

    struct ComPort // custom struct with our desired values
    {
      public string name;
      public string vid;
      public string pid;
      public string description;
    }

    public enum DeviceCommands { MuteOrUnmuteTS, MuteOrUnmuteDC, DownloadTrack, PCShutdown, ClipCapture, SilentMode, GlobalMute, PlayPause, SkipSong, NoCmd };

    // Buffer for Serial Data
    private static string dataBuffer = string.Empty;

    // Serial Port Handle
    private static SerialPort serialPort;

    private static void ToggleSoundMute()
    {
      MainSound.AudioEndpointVolume.Mute = !MainSound.AudioEndpointVolume.Mute;
    }

    private static void ToggleMicMute()
    {
      MainMic.AudioEndpointVolume.Mute = !MainMic.AudioEndpointVolume.Mute;
    }

    // Function to shutdown PC (created by Werti1304)
    private static void Goodbye()
    {
      Process.Start("shutdown", "/s /t 0");
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
          {
            c.vid = mVID.Groups[1].Value;
          }

          if (mPID.Success)
          {
            c.pid = mPID.Groups[1].Value;
          }

          return c;

        }).ToList();
      }
    }

    private static string GetPortName()
    {
      List<ComPort> ports = GetSerialPorts();
      ComPort com = ports.FindLast(c => c.vid.Equals("2341") && c.pid.Equals("0001"));

      return com.name;
    }

    private static DeviceCommands GetCommand(string data)
    {
      if (data.IndexOf("mtntmts") != -1)
      {
        return DeviceCommands.MuteOrUnmuteTS;
      }
      else if (data.IndexOf("mtntmdc") != -1)
      {
        return DeviceCommands.MuteOrUnmuteDC;
      }
      else if (data.IndexOf("dwnldtrck") != -1)
      {
        return DeviceCommands.DownloadTrack;
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
      else if (data.IndexOf("glblmt") != -1)
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

    private async void onSerialDataRecieve(object sender, SerialDataReceivedEventArgs args)
    {
      string strData = serialPort.ReadExisting();
      dataBuffer += strData;

      if (!dataBuffer.EndsWith("\r\n"))
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
        case DeviceCommands.DownloadTrack:
          await DownloadCurrentSongAsync();
          break;
        case DeviceCommands.PCShutdown:
          this.BeginInvoke(() => { Goodbye(); });
          break;
        case DeviceCommands.ClipCapture:
          SendKeys.SendWait("%{F10}");
          break;
        case DeviceCommands.SilentMode:
          this.BeginInvoke(() => { ToggleSoundMute(); });
          break;
        case DeviceCommands.GlobalMute:
          this.BeginInvoke(() => { ToggleMicMute(); });
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

    private async Task AuthServerMain()
    {
      spotifyAuthServer = new EmbedIOAuthServer(new Uri("http://localhost:5000/callback"), 5000);
      await spotifyAuthServer.Start();
    }

    private static async Task DownloadCurrentSongAsync()
    {
      // Refresh Spotify Access Token if needed
      if (spotifyAuthResponse.IsExpired)
      {
        spotifyAuthResponse = await new OAuthClient().RequestToken(new AuthorizationCodeRefreshRequest(spotifyApiCredentials.AppId, spotifyApiCredentials.AppSecret, spotifyApiCredentials.Token));
        spotifyApi = new SpotifyClient(spotifyAuthResponse.AccessToken);
      }

      // Get Track info
      var track = await spotifyApi.Player.GetCurrentlyPlaying(new PlayerCurrentlyPlayingRequest { Market = "DE" });

      if (track == null)
      {
        MessageBox.Show("Error", "Could not get currently played track!");
        return;
      }

      string artist = ((FullTrack)track.Item).Artists[0].Name;
      string trackName = ((FullTrack)track.Item).Name;
      string album = ((FullTrack)track.Item).Album.Name;

      // Get YT Video results
      var searchListRequest = ytService.Search.List("snippet");
      searchListRequest.Q = string.Format("{0} {1} topic", artist, trackName);
      searchListRequest.MaxResults = 3;

      var searchListResponse = await searchListRequest.ExecuteAsync();

      List<string> videos = new List<string>();
      foreach (var searchResult in searchListResponse.Items)
      {
        switch (searchResult.Id.Kind)
        {
          case "youtube#video":
            videos.Add(String.Format(@"https://www.youtube.com/watch?v={0}", searchResult.Id.VideoId));
            break;
        }
      }

      string musicRoot = Path.GetFullPath("C:\\Users\\lukas\\Desktop\\Music");
      string albumPath = Path.Combine(musicRoot, artist, album);

      if (!Directory.Exists(albumPath))
      {
        Directory.CreateDirectory(albumPath);
      }

      // Download Audio File
      ProcessStartInfo info = new ProcessStartInfo("CMD.exe");
      info.Arguments = String.Format
      (
        "/C {0}\\yt-dlp.exe -f 140 {1} -o \"{2}\\{3}.m4a\"", 
        Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), 
        videos[0],
        albumPath,
        trackName
      );
      Process dwnldPrcss = Process.Start(info);
      dwnldPrcss.WaitForExit();

      // Edit ID3-Tags of Audio File
      var tagFile = TagLib.File.Create(string.Format("{0}\\{1}.m4a", albumPath, trackName));
      tagFile.Tag.Album = album;
      tagFile.Tag.Performers = new string[] { artist };
      tagFile.Tag.Title = trackName;
      tagFile.Save();
    }

    public MainForm()
    {
      InitializeComponent();
    }

    private async void MainForm_Load(object sender, EventArgs e)
    {
      // Check for USB Serial Device
      string portName;
      while ((portName = GetPortName()) == null)
      {
      }

      serialPort = new SerialPort(portName, 9600);
      serialPort.DataReceived += new SerialDataReceivedEventHandler(onSerialDataRecieve);

      // Get Handles for Mic and main Sound device
      var enumerator = new MMDeviceEnumerator();
      MainSound = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
      MainMic = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications);

      // Get API Credentials
      string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "apiCredentials.json");
      if(!File.Exists(filePath))
      {
        MessageBox.Show("File Not Found", "API Credentials file not found!");
        Environment.Exit(-1);
      }

      string fileContent = string.Empty;
      using (StreamReader streamReader = new StreamReader(filePath))
      {
        fileContent = streamReader.ReadToEnd();
      }

      JsonRoot jsonRoot = JsonConvert.DeserializeObject<JsonRoot>(fileContent);

      if(jsonRoot == null)
      {
        MessageBox.Show("API Credentials Error", "Couldn't read API Credentials!");
        Environment.Exit(-1);
      }

      // Get Youtube API Access
      ytApiCredentials = jsonRoot.ApiCredentials.Where(x => x.ApiName == "YouTubeApi").Single();
      ytService = new YouTubeService(new Google.Apis.Services.BaseClientService.Initializer
      {
        ApiKey = ytApiCredentials.Token,
        ApplicationName = this.GetType().ToString()
      });

      // Get Spotify API Access
      spotifyApiCredentials = jsonRoot.ApiCredentials.Where(x => x.ApiName == "SpotifyWebApi").Single();
      await AuthServerMain();
      spotifyAuthResponse = await new OAuthClient().RequestToken(new AuthorizationCodeRefreshRequest(spotifyApiCredentials.AppId, spotifyApiCredentials.AppSecret, spotifyApiCredentials.Token));
      spotifyApi = new SpotifyClient(spotifyAuthResponse.AccessToken);

      // Try to start Serial Connection
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