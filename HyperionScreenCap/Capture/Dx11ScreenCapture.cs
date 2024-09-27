using HyperionScreenCap.Capture;
using SharpDX;
using SharpDX.Direct3D11;
using SharpDX.DXGI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using log4net;

namespace HyperionScreenCap
{
    class DX11ScreenCapture : IScreenCapture
    {
        private int _adapterIndex;
        private int _monitorIndex;
        private int _scalingFactor;
        private int _maxFps;
        private int _frameCaptureTimeout;

        private Factory1 _factory;
        private Adapter _adapter;
        private Output _output;
        private Output1 _output1;
        private SharpDX.Direct3D11.Device _device;
        private Texture2D _stagingTexture;
        private Texture2D _smallerTexture;
        private ShaderResourceView _smallerTextureView;
        private OutputDuplication _duplicatedOutput;
        private int _scalingFactorLog2;
        private int _width;
        private int _height;
        private byte[] _lastCapturedFrame;
        private int _minCaptureTime;
        private Stopwatch _captureTimer;
        private bool _desktopDuplicatorInvalid;
        private bool _disposed;

        private int _reinitializationAttempts = 0;
        private const int MAX_REINITIALIZATION_ATTEMPTS = 5;
        private const int REINITIALIZATION_DELAY_MS = 1000;

        public int CaptureWidth { get; private set; }
        public int CaptureHeight { get; private set; }

        private static readonly ILog LOG = LogManager.GetLogger(typeof(DX11ScreenCapture));

        public static String GetAvailableMonitors()
        {
            StringBuilder response = new StringBuilder();
            using ( Factory1 factory = new Factory1() )
            {
                int adapterIndex = 0;
                foreach(Adapter adapter in factory.Adapters)
                {
                    response.Append($"Adapter Index {adapterIndex++}: {adapter.Description.Description}\n");
                    int outputIndex = 0;
                    foreach(Output output in adapter.Outputs)
                    {
                        response.Append($"\tMonitor Index {outputIndex++}: {output.Description.DeviceName}");
                        var desktopBounds = output.Description.DesktopBounds;
                        response.Append($" {desktopBounds.Right - desktopBounds.Left}×{desktopBounds.Bottom - desktopBounds.Top}\n");
                    }
                    response.Append("\n");
                }
            }
            return response.ToString();
        }

        public DX11ScreenCapture(int adapterIndex, int monitorIndex, int scalingFactor, int maxFps, int frameCaptureTimeout)
        {
            _adapterIndex = adapterIndex;
            _monitorIndex = monitorIndex;
            _scalingFactor = scalingFactor;
            _maxFps = maxFps;
            _frameCaptureTimeout = frameCaptureTimeout;
            _disposed = true;
        }

        public void Initialize()
        {
            Dispose();

            int retryCount = 0;
            const int maxRetries = 5;
            const int retryDelay = 1000; // 1 second

            while (retryCount < maxRetries)
            {
                try
                {
                    InitializeInternal();
                    return; // If successful, exit the method
                }
                catch (SharpDXException ex)
                {
                    LOG.Error($"SharpDX exception during initialization (attempt {retryCount + 1}/{maxRetries}): {ex.Message}");
                    retryCount++;
                    if (retryCount >= maxRetries)
                    {
                        throw new Exception("Failed to initialize DX11 screen capture after multiple attempts", ex);
                    }
                    Thread.Sleep(retryDelay);
                }
                catch (Exception ex)
                {
                    LOG.Error($"Unexpected exception during initialization: {ex.Message}");
                    throw;
                }
            }
            InitDesktopDuplicator();
        }

        private void InitializeInternal()
        {
            // Move all the existing initialization code here
            int mipLevels;
            if (_scalingFactor == 1)
                mipLevels = 1;
            else if (_scalingFactor > 0 && _scalingFactor % 2 == 0)
            {
                _scalingFactorLog2 = Convert.ToInt32(Math.Log(_scalingFactor, 2));
                mipLevels = 2 + _scalingFactorLog2 - 1;
            }
            else
                throw new Exception("Invalid scaling factor. Allowed values are 1, 2, 4, etc.");

            _factory?.Dispose();
            _adapter?.Dispose();
            _output?.Dispose();
            _output1?.Dispose();
            _device?.Dispose();
            _stagingTexture?.Dispose();
            _smallerTexture?.Dispose();
            _smallerTextureView?.Dispose();

            _factory = new Factory1();
            _adapter = _factory.GetAdapter1(_adapterIndex);

            _device = new SharpDX.Direct3D11.Device(_adapter);

            _output = _adapter.GetOutput(_monitorIndex);
            if (_output == null)
            {
                throw new Exception($"No output found for adapter {_adapterIndex} and monitor {_monitorIndex}");
            }
            _output1 = _output.QueryInterface<Output1>();

            var desktopBounds = _output.Description.DesktopBounds;
            _width = desktopBounds.Right - desktopBounds.Left;
            _height = desktopBounds.Bottom - desktopBounds.Top;

            CaptureWidth = _width / _scalingFactor;
            CaptureHeight = _height / _scalingFactor;

            var stagingTextureDesc = new Texture2DDescription
            {
                CpuAccessFlags = CpuAccessFlags.Read,
                BindFlags = BindFlags.None,
                Format = Format.B8G8R8A8_UNorm,
                Width = CaptureWidth,
                Height = CaptureHeight,
                OptionFlags = ResourceOptionFlags.None,
                MipLevels = 1,
                ArraySize = 1,
                SampleDescription = { Count = 1, Quality = 0 },
                Usage = ResourceUsage.Staging
            };
            _stagingTexture = new Texture2D(_device, stagingTextureDesc);

            var smallerTextureDesc = new Texture2DDescription
            {
                CpuAccessFlags = CpuAccessFlags.None,
                BindFlags = BindFlags.RenderTarget | BindFlags.ShaderResource,
                Format = Format.B8G8R8A8_UNorm,
                Width = _width,
                Height = _height,
                OptionFlags = ResourceOptionFlags.GenerateMipMaps,
                MipLevels = mipLevels,
                ArraySize = 1,
                SampleDescription = { Count = 1, Quality = 0 },
                Usage = ResourceUsage.Default
            };
            _smallerTexture = new Texture2D(_device, smallerTextureDesc);
            _smallerTextureView = new ShaderResourceView(_device, _smallerTexture);

            _minCaptureTime = 1000 / _maxFps;
            _captureTimer = new Stopwatch();
            _disposed = false;

            InitDesktopDuplicator();
        }

        private void InitDesktopDuplicator()
        {
            try
            {
                _duplicatedOutput?.Dispose();
                _duplicatedOutput = _output1.DuplicateOutput(_device);
                _desktopDuplicatorInvalid = false;
                LOG.Debug("Desktop duplicator initialized successfully.");
            }
            catch (SharpDXException ex)
            {
                LOG.Error($"SharpDXException in InitDesktopDuplicator: {ex.Message}", ex);
                _desktopDuplicatorInvalid = true;
                throw;
            }
        }

        private void ReinitializeOutputAndDevice()
        {
            LOG.Debug("Reinitializing Output and Device...");
            try
            {
                _factory?.Dispose();
                _adapter?.Dispose();
                _output?.Dispose();
                _output1?.Dispose();
                _device?.Dispose();

                _factory = new Factory1();
                if (_factory == null) throw new InvalidOperationException("Failed to create Factory1");

                _adapter = _factory.GetAdapter1(_adapterIndex);
                if (_adapter == null) throw new InvalidOperationException($"No adapter found for index {_adapterIndex}");

                _device = new SharpDX.Direct3D11.Device(_adapter);
                if (_device == null) throw new InvalidOperationException("Failed to create Device");

                _output = _adapter.GetOutput(_monitorIndex);
                if (_output == null) throw new InvalidOperationException($"No output found for adapter {_adapterIndex} and monitor {_monitorIndex}");

                _output1 = _output.QueryInterface<Output1>();
                if (_output1 == null) throw new InvalidOperationException("Failed to query Output1 interface");

                LOG.Info("Successfully reinitialized Output and Device.");
            }
            catch (Exception ex)
            {
                LOG.Error($"Failed to reinitialize Output and Device: {ex.Message}", ex);
                throw;
            }
        }

        private void EnsureDeviceInitialized()
        {
            if (_device == null || _device.IsDisposed)
            {
                LOG.Warn("Device is null or disposed. Attempting to reinitialize...");
                try
                {
                    ReinitializeOutputAndDevice();
                }
                catch (Exception ex)
                {
                    LOG.Error($"Failed to reinitialize device: {ex.Message}", ex);
                    throw new InvalidOperationException("Failed to reinitialize device", ex);
                }
            }
        }

        public byte[] Capture()
        {
            if (_desktopDuplicatorInvalid)
            {
                LOG.Warn("Desktop duplicator is invalid. Attempting to reinitialize...");
                if (!ReinitializeDesktopDuplicator())
                {
                    throw new InvalidOperationException("Failed to reinitialize desktop duplicator after multiple attempts");
                }
            }

            _captureTimer.Restart();
            try
            {
                byte[] response = ManagedCapture();
                _captureTimer.Stop();
                _reinitializationAttempts = 0; // Reset the counter on successful capture
                return response;
            }
            catch (InvalidOperationException ex)
            {
                LOG.Error($"Capture failed: {ex.Message}", ex);
                _desktopDuplicatorInvalid = true;
                throw;
            }
        }

        private bool ReinitializeDesktopDuplicator()
        {
            while (_reinitializationAttempts < MAX_REINITIALIZATION_ATTEMPTS)
            {
                try
                {
                    _reinitializationAttempts++;
                    LOG.Info($"Reinitialization attempt {_reinitializationAttempts} of {MAX_REINITIALIZATION_ATTEMPTS}");

                    // Only reinitialize the desktop duplicator
                    InitDesktopDuplicator();

                    _desktopDuplicatorInvalid = false;
                    LOG.Info("Desktop duplicator reinitialized successfully.");
                    return true;
                }
                catch (Exception ex)
                {
                    LOG.Error($"Failed to reinitialize desktop duplicator: {ex.Message}", ex);
                    Thread.Sleep(REINITIALIZATION_DELAY_MS * _reinitializationAttempts);
                }
            }
            return false;
        }

        private byte[] ManagedCapture()
        {
            SharpDX.DXGI.Resource screenResource = null;
            OutputDuplicateFrameInformation duplicateFrameInformation;

            try
            {
                EnsureDeviceInitialized();

                try
                {
                    // Try to get duplicated frame within given time
                    _duplicatedOutput.AcquireNextFrame(_frameCaptureTimeout, out duplicateFrameInformation, out screenResource);

                    if (duplicateFrameInformation.LastPresentTime == 0 && _lastCapturedFrame != null)
                        return _lastCapturedFrame;
                }
                catch (SharpDXException ex)
                {
                    if (ex.ResultCode.Code == SharpDX.DXGI.ResultCode.WaitTimeout.Code && _lastCapturedFrame != null)
                        return _lastCapturedFrame;

                    if (ex.ResultCode.Code == SharpDX.DXGI.ResultCode.AccessLost.Code)
                    {
                        _desktopDuplicatorInvalid = true;
                        throw new InvalidOperationException("Desktop duplicator access lost", ex);
                    }

                    throw;
                }

                // Check if scaling is used
                if (CaptureWidth != _width)
                {
                    // Copy resource into memory that can be accessed by the CPU
                    using (var screenTexture2D = screenResource.QueryInterface<Texture2D>())
                    {
                        if (_device == null || _device.ImmediateContext == null)
                        {
                            throw new InvalidOperationException("Device or ImmediateContext is null");
                        }
                        _device.ImmediateContext.CopySubresourceRegion(screenTexture2D, 0, null, _smallerTexture, 0);
                    }

                    // Generates the mipmap of the screen
                    _device.ImmediateContext.GenerateMips(_smallerTextureView);

                    // Copy the mipmap of smallerTexture (size/ scalingFactor) to the staging texture: 1 for /2, 2 for /4...etc
                    _device.ImmediateContext.CopySubresourceRegion(_smallerTexture, _scalingFactorLog2, null, _stagingTexture, 0);
                }
                else
                {
                    // Copy resource into memory that can be accessed by the CPU
                    using (var screenTexture2D = screenResource.QueryInterface<Texture2D>())
                    {
                        if (_device == null || _device.ImmediateContext == null)
                        {
                            throw new InvalidOperationException("Device or ImmediateContext is null");
                        }
                        _device.ImmediateContext.CopyResource(screenTexture2D, _stagingTexture);
                    }
                }

                // Get the desktop capture texture
                var mapSource = _device.ImmediateContext.MapSubresource(_stagingTexture, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);
                _lastCapturedFrame = ToRGBArray(mapSource);
                return _lastCapturedFrame;
            }
            catch (Exception ex)
            {
                LOG.Error($"Error in ManagedCapture: {ex.Message}", ex);
                throw;
            }
            finally
            {
                screenResource?.Dispose();
                // Fixed OUT_OF_MEMORY issue on AMD Radeon cards. Ignoring all exceptions during unmapping.
                try { _device?.ImmediateContext?.UnmapSubresource(_stagingTexture, 0); } catch { };
                // Ignore DXGI_ERROR_INVALID_CALL, DXGI_ERROR_ACCESS_LOST errors since capture is already complete
                try { _duplicatedOutput?.ReleaseFrame(); } catch { }
            }
        }

        /// <summary>
        /// Reads from the memory locations pointed to by the DataBox and saves it into a byte array
        /// ignoring the alpha component of each pixel.
        /// </summary>
        /// <param name="mapSource"></param>
        /// <returns></returns>
        private byte[] ToRGBArray(DataBox mapSource)
        {
            var sourcePtr = mapSource.DataPointer;
            byte[] bytes = new byte[CaptureWidth * 3 * CaptureHeight];
            int byteIndex = 0;
            for ( int y = 0; y < CaptureHeight; y++ )
            {
                Int32[] rowData = new Int32[CaptureWidth];
                Marshal.Copy(sourcePtr, rowData, 0, CaptureWidth);

                foreach ( Int32 pixelData in rowData )
                {
                    byte[] values = BitConverter.GetBytes(pixelData);
                    if ( BitConverter.IsLittleEndian )
                    {
                        // Byte order : bgra
                        bytes[byteIndex++] = values[2];
                        bytes[byteIndex++] = values[1];
                        bytes[byteIndex++] = values[0];
                    }
                    else
                    {
                        // Byte order : argb
                        bytes[byteIndex++] = values[1];
                        bytes[byteIndex++] = values[2];
                        bytes[byteIndex++] = values[3];
                    }
                }

                sourcePtr = IntPtr.Add(sourcePtr, mapSource.RowPitch);
            }
            return bytes;
        }

        public void DelayNextCapture()
        {
            int remainingFrameTime = _minCaptureTime - (int)_captureTimer.ElapsedMilliseconds;
            if ( remainingFrameTime > 0 )
            {
                Thread.Sleep(remainingFrameTime);
            }
        }

        public void Dispose()
        {
            _duplicatedOutput?.Dispose();
            _output1?.Dispose();
            _output?.Dispose();
            _stagingTexture?.Dispose();
            _smallerTexture?.Dispose();
            _smallerTextureView?.Dispose();
            _device?.Dispose();
            _adapter?.Dispose();
            _factory?.Dispose();
            _lastCapturedFrame = null;
            _disposed = true;
        }

        public bool IsDisposed()
        {
            return _disposed;
        }
    }
}
