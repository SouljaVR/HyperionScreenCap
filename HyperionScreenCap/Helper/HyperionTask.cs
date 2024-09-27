using HyperionScreenCap.Capture;
using HyperionScreenCap.Config;
using HyperionScreenCap.Model;
using HyperionScreenCap.Networking;
using log4net;
using System;
using System.Collections.Generic;
using System.Threading;

namespace HyperionScreenCap.Helper
{
    class HyperionTask // TODO: Remove notifications from here
    {
        private static readonly ILog LOG = LogManager.GetLogger(typeof(HyperionTask));

        private HyperionTaskConfiguration _configuration;
        private NotificationUtils _notificationUtils;

        private IScreenCapture _screenCapture;
        private List<HyperionClient> _hyperionClients;
        public bool CaptureEnabled { get; private set; }
        private Thread _captureThread;

        public event EventHandler OnCaptureDisabled;

        public HyperionTask(HyperionTaskConfiguration configuration, NotificationUtils notificationUtils)
        {
            this._configuration = configuration;
            this._notificationUtils = notificationUtils;
            this._hyperionClients = new List<HyperionClient>();
        }

        private void InitScreenCapture()
        {
            if ( _screenCapture != null && !_screenCapture.IsDisposed() )
            {
                // Screen capture already initialized. Ignoring request.
                return;
            }
            try
            {
                LOG.Info($"{this}: Initializing screen capture");
                _screenCapture.Initialize();
                LOG.Info($"{this}: Screen capture initialized");
            }
            catch ( Exception ex )
            {
                _screenCapture?.Dispose();
                throw new Exception("Failed to initialize screen capture: " + ex.Message, ex);
            }
        }

        private String GetHyperionInitFailedMsg(HyperionClient hyperionClient)
        {
            return $"Failed to connect to Hyperion server using {hyperionClient}";
        }

        private void InstantiateHyperionClients()
        {
            foreach ( HyperionServer server in _configuration.HyperionServers )
            {
                switch (server.Protocol)
                {
                    case HyperionServerProtocol.PROTOCOL_BUFFERS:
                        _hyperionClients.Add(new ProtoClient(server.Host, server.Port, server.Priority, server.MessageDuration));
                        break;

                    case HyperionServerProtocol.FLAT_BUFFERS:
                        _hyperionClients.Add(new FbsClinet(server.Host, server.Port, server.Priority, server.MessageDuration));
                        break;

                    default:
                        throw new NotImplementedException($"Hyperion server protocol {server.Protocol} is not supported yet");
                }
                
            }
        }

        private void InstantiateScreenCapture()
        {
            switch ( _configuration.CaptureMethod )
            {
                case CaptureMethod.DX9:
                    _screenCapture = new DX9ScreenCapture(_configuration.Dx9MonitorIndex, _configuration.Dx9CaptureWidth, _configuration.Dx9CaptureHeight,
                        _configuration.Dx9CaptureInterval);
                    break;

                case CaptureMethod.DX11:
                    _screenCapture = new DX11ScreenCapture(_configuration.Dx11AdapterIndex, _configuration.Dx11MonitorIndex, _configuration.Dx11ImageScalingFactor,
                        _configuration.Dx11MaxFps, _configuration.Dx11FrameCaptureTimeout);
                    break;

                default:
                    throw new NotImplementedException($"The capture method {_configuration.CaptureMethod} is not supported yet");
            }
        }

        private void DisposeHyperionClients()
        {
            foreach ( HyperionClient hyperionClient in _hyperionClients )
            {
                hyperionClient?.Dispose();
            }
        }

        private void ConnectHyperionClients()
        {
            foreach (HyperionClient hyperionClient in _hyperionClients)
            {
                try
                {
                    LOG.Info($"{this}: Connecting {hyperionClient}");
                    hyperionClient.Dispose(); // Ensure any existing connection is closed
                    hyperionClient.Connect(); // This will now send registration
                    if (hyperionClient.IsConnected())
                    {
                        LOG.Info($"{this}: {hyperionClient} connected");
                        _notificationUtils.Info($"Connected to Hyperion server using {hyperionClient}!");

                        // Only send initial frame if screen capture is initialized
                        if (_screenCapture != null)
                        {
                            hyperionClient.SendInitialFrame(_screenCapture.CaptureWidth, _screenCapture.CaptureHeight);

                            // Send an actual captured frame immediately
                            byte[] initialFrame = CaptureInitialFrame();
                            hyperionClient.SendImageData(initialFrame, _screenCapture.CaptureWidth, _screenCapture.CaptureHeight);
                        }
                    }
                    else
                    {
                        throw new Exception(GetHyperionInitFailedMsg(hyperionClient));
                    }
                }
                catch (Exception ex)
                {
                    LOG.Error($"{this}: Failed to connect to Hyperion server: {ex.Message}", ex);
                    throw;
                }
            }
        }

        private byte[] CaptureInitialFrame()
        {
            try
            {
                return _screenCapture.Capture();
            }
            catch (Exception ex)
            {
                LOG.Error($"{this}: Failed to capture initial frame: {ex.Message}", ex);
                // Return a black frame as a fallback
                return new byte[_screenCapture.CaptureWidth * _screenCapture.CaptureHeight * 3];
            }
        }

        private void TransmitNextFrame()
        {
            foreach ( HyperionClient hyperionClient in _hyperionClients )
            {
                try
                {
                    byte[] imageData = _screenCapture.Capture();
                    hyperionClient.SendImageData(imageData, _screenCapture.CaptureWidth, _screenCapture.CaptureHeight);

                    // Uncomment the following to enable debugging
                    // MiscUtils.SaveRGBArrayToImageFile(imageData, _screenCapture.CaptureWidth, _screenCapture.CaptureHeight, AppConstants.DEBUG_IMAGE_FILE_NAME);
                }
                catch ( Exception ex )
                {
                    throw new Exception("Error occured while sending image to server: " + ex.Message, ex);
                }
            }
        }

        public void RestartCapture()
        {
            DisableCapture();
            Thread.Sleep(1000); // Wait a bit before restarting
            EnableCapture();
        }

        private void StartCapture()
        {
            InstantiateScreenCapture();
            InstantiateHyperionClients();
            int captureAttempt = 1;
            while (CaptureEnabled)
            {
                try
                {
                    InitScreenCapture();
                    ConnectHyperionClients();
                    while (CaptureEnabled)
                    {
                        TransmitNextFrame();
                        _screenCapture.DelayNextCapture();
                    }
                }
                catch (Exception ex)
                {
                    LOG.Error($"{this}: Exception in screen capture attempt: {captureAttempt}", ex);
                    if (captureAttempt > AppConstants.REINIT_CAPTURE_AFTER_ATTEMPTS)
                    {
                        LOG.Info($"{this}: Attempting to recreate screen capture object and reconnect to Hyperion");
                        RecreateScreenCaptureAndReconnect();
                        return; // Exit this method, as a new capture thread has been started
                    }
                    Thread.Sleep(AppConstants.CAPTURE_FAILED_COOLDOWN_MILLIS);
                    captureAttempt++;

                    if (captureAttempt > AppConstants.MAX_CAPTURE_ATTEMPTS)
                    {
                        LOG.Error($"{this}: Max screen capture attempts reached. Disabling capture.");
                        CaptureEnabled = false;
                        OnCaptureDisabled?.Invoke(this, EventArgs.Empty);
                    }
                }
            }
        }

        private void RecreateScreenCaptureAndReconnect()
        {
            LOG.Info($"{this}: Recreating screen capture and reconnecting to Hyperion");

            // Dispose of existing resources
            _screenCapture?.Dispose();
            _screenCapture = null;

            foreach (var client in _hyperionClients)
            {
                client.Dispose();
            }
            _hyperionClients.Clear();

            // Wait a bit before reconnecting
            Thread.Sleep(2000);

            // Let StartCapture handle the reinitialization
        }

        private void TryStartCapture()
        {
            try // Properly dispose everything object when turning off capture
            {
                StartCapture();
            }
            finally
            {
                _screenCapture?.Dispose();
                DisposeHyperionClients();
            }
            LOG.Info($"{this}: Screen Capture finished");
        }

        public void EnableCapture()
        {
            LOG.Info($"{this}: Enabling screen capture");
            if (_captureThread == null || !_captureThread.IsAlive)
            {
                CaptureEnabled = true;
                _captureThread = new Thread(TryStartCapture) { IsBackground = true };
                _captureThread.Start();
            }
            else
            {
                LOG.Warn($"{this}: Capture thread is already running");
            }
        }


        public void DisableCapture()
        {
            LOG.Info($"{this}: Disabling screen capture");
            CaptureEnabled = false;
        }

        public override String ToString()
        {
            return $"HyperionTask[ConfigurationId: {_configuration.Id}]";
        }
    }
}
