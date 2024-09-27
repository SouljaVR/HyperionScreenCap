using HyperionScreenCap.Capture;
using SharpDX;
using SharpDX.Direct3D11;
using SharpDX.DXGI;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using log4net;

namespace HyperionScreenCap
{
    // Custom exception to indicate that capture is paused
    public class CapturePausedException : Exception
    {
        public CapturePausedException(string message)
            : base(message)
        {
        }
    }

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
        private bool _disposed;

        public int CaptureWidth { get; private set; }
        public int CaptureHeight { get; private set; }

        private static readonly ILog LOG = LogManager.GetLogger(typeof(DX11ScreenCapture));

        public static String GetAvailableMonitors()
        {
            StringBuilder response = new StringBuilder();
            using (Factory1 factory = new Factory1())
            {
                int adapterIndex = 0;
                foreach (Adapter adapter in factory.Adapters)
                {
                    response.Append($"Adapter Index {adapterIndex++}: {adapter.Description.Description}\n");
                    int outputIndex = 0;
                    foreach (Output output in adapter.Outputs)
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
            Dispose(); // Ensure previous resources are released

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

            // Create DXGI Factory1
            _factory = new Factory1();
            _adapter = _factory.GetAdapter1(_adapterIndex);

            // Create device from Adapter
            _device = new SharpDX.Direct3D11.Device(_adapter);

            // Get DXGI.Output
            _output = _adapter.GetOutput(_monitorIndex);
            if (_output == null)
            {
                throw new Exception($"Failed to get output for monitor index {_monitorIndex}.");
            }

            _output1 = _output.QueryInterface<Output1>();
            if (_output1 == null)
            {
                throw new Exception("Failed to get Output1 interface.");
            }

            // Width/Height of desktop to capture
            var desktopBounds = _output.Description.DesktopBounds;
            _width = desktopBounds.Right - desktopBounds.Left;
            _height = desktopBounds.Bottom - desktopBounds.Top;

            CaptureWidth = _width / _scalingFactor;
            CaptureHeight = _height / _scalingFactor;

            // Create Staging texture CPU-accessible
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

            // Create smaller texture to downscale the captured image
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
                _duplicatedOutput = null;

                if (_output1 == null || _device == null)
                {
                    throw new Exception("Output1 or Device is null. Cannot initialize desktop duplicator.");
                }

                _duplicatedOutput = _output1.DuplicateOutput(_device);
                LOG.Debug("Desktop duplicator initialized successfully.");
            }
            catch (SharpDXException ex)
            {
                LOG.Error($"SharpDXException in InitDesktopDuplicator: {ex.Message}", ex);
                throw;
            }
        }

        public byte[] Capture()
        {
            _captureTimer.Restart();
            try
            {
                byte[] response = ManagedCapture();
                _captureTimer.Stop();
                return response;
            }
            catch (SharpDXException ex)
            {
                if (ex.ResultCode == SharpDX.DXGI.ResultCode.AccessLost ||
                    ex.ResultCode == SharpDX.DXGI.ResultCode.AccessDenied ||
                    ex.ResultCode == SharpDX.DXGI.ResultCode.SessionDisconnected ||
                    ex.ResultCode == SharpDX.DXGI.ResultCode.DeviceRemoved ||
                    ex.ResultCode == SharpDX.DXGI.ResultCode.InvalidCall)
                {
                    LOG.Warn($"Capture failed due to device loss: {ex.Message}. Attempting to reinitialize.");
                    Reinitialize();
                    throw new CapturePausedException("Capture is paused due to device loss.");
                }
                else
                {
                    LOG.Error($"Capture failed: {ex.Message}", ex);
                    throw;
                }
            }
            catch (Exception ex)
            {
                LOG.Error($"Capture failed: {ex.Message}", ex);
                throw;
            }
        }

        private void Reinitialize()
        {
            Dispose();
            int attempt = 0;
            const int maxAttempts = 10;
            const int delayBetweenAttempts = 1000; // milliseconds

            while (attempt < maxAttempts)
            {
                try
                {
                    attempt++;
                    Initialize();
                    LOG.Info("Reinitialized capture successfully.");
                    return;
                }
                catch (Exception ex)
                {
                    LOG.Error($"Failed to reinitialize capture (attempt {attempt}/{maxAttempts}): {ex.Message}", ex);
                    Thread.Sleep(delayBetweenAttempts);
                }
            }
            throw new CapturePausedException("Failed to reinitialize capture after multiple attempts.");
        }

        private byte[] ManagedCapture()
        {
            if (_duplicatedOutput == null)
            {
                throw new Exception("DuplicatedOutput is null. Cannot capture.");
            }

            SharpDX.DXGI.Resource screenResource = null;
            OutputDuplicateFrameInformation duplicateFrameInformation;

            try
            {
                // Try to get duplicated frame within given time
                _duplicatedOutput.AcquireNextFrame(_frameCaptureTimeout, out duplicateFrameInformation, out screenResource);

                if (duplicateFrameInformation.LastPresentTime == 0 && _lastCapturedFrame != null)
                    return _lastCapturedFrame;

                // Check if scaling is used
                if (CaptureWidth != _width)
                {
                    // Copy resource into memory that can be accessed by the CPU
                    using (var screenTexture2D = screenResource.QueryInterface<Texture2D>())
                    {
                        _device.ImmediateContext.CopySubresourceRegion(screenTexture2D, 0, null, _smallerTexture, 0);
                    }

                    // Generates the mipmap of the screen
                    _device.ImmediateContext.GenerateMips(_smallerTextureView);

                    // Copy the mipmap of smallerTexture (size/ scalingFactor) to the staging texture
                    _device.ImmediateContext.CopySubresourceRegion(_smallerTexture, _scalingFactorLog2, null, _stagingTexture, 0);
                }
                else
                {
                    // Copy resource into memory that can be accessed by the CPU
                    using (var screenTexture2D = screenResource.QueryInterface<Texture2D>())
                    {
                        _device.ImmediateContext.CopyResource(screenTexture2D, _stagingTexture);
                    }
                }

                // Get the desktop capture texture
                var mapSource = _device.ImmediateContext.MapSubresource(_stagingTexture, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);
                _lastCapturedFrame = ToRGBArray(mapSource);
                return _lastCapturedFrame;
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
            for (int y = 0; y < CaptureHeight; y++)
            {
                Int32[] rowData = new Int32[CaptureWidth];
                Marshal.Copy(sourcePtr, rowData, 0, CaptureWidth);

                foreach (Int32 pixelData in rowData)
                {
                    byte[] values = BitConverter.GetBytes(pixelData);
                    if (BitConverter.IsLittleEndian)
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
            if (remainingFrameTime > 0)
            {
                Thread.Sleep(remainingFrameTime);
            }
        }

        public void Dispose()
        {
            try { _duplicatedOutput?.Dispose(); } catch (Exception ex) { LOG.Error("Exception during Dispose of _duplicatedOutput", ex); }
            _duplicatedOutput = null;

            try { _output1?.Dispose(); } catch (Exception ex) { LOG.Error("Exception during Dispose of _output1", ex); }
            _output1 = null;

            try { _output?.Dispose(); } catch (Exception ex) { LOG.Error("Exception during Dispose of _output", ex); }
            _output = null;

            try { _stagingTexture?.Dispose(); } catch (Exception ex) { LOG.Error("Exception during Dispose of _stagingTexture", ex); }
            _stagingTexture = null;

            try { _smallerTexture?.Dispose(); } catch (Exception ex) { LOG.Error("Exception during Dispose of _smallerTexture", ex); }
            _smallerTexture = null;

            try { _smallerTextureView?.Dispose(); } catch (Exception ex) { LOG.Error("Exception during Dispose of _smallerTextureView", ex); }
            _smallerTextureView = null;

            try { _device?.Dispose(); } catch (Exception ex) { LOG.Error("Exception during Dispose of _device", ex); }
            _device = null;

            try { _adapter?.Dispose(); } catch (Exception ex) { LOG.Error("Exception during Dispose of _adapter", ex); }
            _adapter = null;

            try { _factory?.Dispose(); } catch (Exception ex) { LOG.Error("Exception during Dispose of _factory", ex); }
            _factory = null;

            _lastCapturedFrame = null;
            _disposed = true;
        }

        public bool IsDisposed()
        {
            return _disposed;
        }
    }
}
