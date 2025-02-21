using System;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;

public class Program {
    // Import the ExtractIconEx method from Shell32.dll to get icons from DLL files
    [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern void ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, IntPtr piSmallVersion, int amountIcons);

    // Splash screen indicating the microphone mute state
    public class SplashScreen : Form {
        public SplashScreen(bool wasMuted)
        {
            // Set the splash screen properties
            FormBorderStyle = FormBorderStyle.None;
            TopMost = true;
            ShowIcon = false;
            ShowInTaskbar = false;

            // Position the splash screen at the bottom center of the screen
            StartPosition = FormStartPosition.Manual;
            Location = new Point((Screen.PrimaryScreen.WorkingArea.Width - Width) / 2,
                                    Screen.PrimaryScreen.WorkingArea.Bottom - Height);

            // Set splash screen background and transparency
            BackgroundImageLayout = ImageLayout.Center;
            BackColor = Color.Black;
            TransparencyKey = Color.Black;

            // Get icons
            IntPtr micIcon = IntPtr.Zero;
            ExtractIconEx(Environment.GetEnvironmentVariable("SystemRoot") + @"\System32\ddores.dll", 3, out micIcon, IntPtr.Zero, 1);
            IntPtr xIcon = IntPtr.Zero;
            ExtractIconEx(Environment.GetEnvironmentVariable("SystemRoot") + @"\System32\imageres.dll", 84, out xIcon, IntPtr.Zero, 1);

            // Show microphone icon
            BackgroundImage = Icon.FromHandle(micIcon).ToBitmap();
            // Overlay with X if microphone is now muted
            if (!wasMuted) {
                Graphics.FromImage(BackgroundImage).DrawImage(
                    Icon.FromHandle(xIcon).ToBitmap(), new Point(0, 0));
            }
        }
        // Don't steal focus
        protected override bool ShowWithoutActivation {
            get { return true; }
        }
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams baseParams = base.CreateParams;
                const int WS_EX_NOACTIVATE = 0x08000000;
                const int WS_EX_TOOLWINDOW = 0x00000080;
                baseParams.ExStyle |= ( int )( WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW );
                return baseParams;
            }
        }
    }

    // Define necessary enumeration and interfaces for microphone control
    enum EDataFlow { eRender, eCapture, eAll }
    enum ERole { eConsole, eMultimedia, eCommunications }

    [ComImport]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDeviceEnumerator {
        void _VtblGap1_1(); // Skip one method
        [PreserveSig]
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDevice {
        [PreserveSig]
        int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, out IAudioEndpointVolume ppInterface);
    }

    [ComImport]
    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioEndpointVolume {
        void _VtblGap1_11(); // Skip eleven methods
        [PreserveSig]
        int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, IntPtr pguidEventContext);
        [PreserveSig]
        int GetMute(out bool pbMute);
    }

    // Method to toggle the microphone mute state
    public static bool ToggleMuteMicrophone() {
        // Create an instance of the device enumerator
        var CLSID_MMDeviceEnumerator = new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E");
        IMMDeviceEnumerator enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(
            Type.GetTypeFromCLSID(CLSID_MMDeviceEnumerator));

        // Get the default audio endpoint for capture
        IMMDevice device;
        enumerator.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia, out device);

        // Activate the audio endpoint volume interface
        var IID_IAudioEndpointVolume = typeof(IAudioEndpointVolume).GUID;
        IAudioEndpointVolume volume;
        device.Activate(ref IID_IAudioEndpointVolume, 23, IntPtr.Zero, out volume);

        // Get the current mute state and toggle it
        bool wasMuted;
        volume.GetMute(out wasMuted);
        volume.SetMute(!wasMuted, IntPtr.Zero);

        return wasMuted;
    }
}