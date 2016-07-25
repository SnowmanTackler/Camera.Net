// Written by Sam Seifert.

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices.ComTypes;

// Microsoft.Win32 is used for SystemEvents namespace
using Microsoft.Win32;

// Use DirectShowLib (LGPL v2.1)
using DirectShowLib;

namespace CameraNet
{
    /// <summary>
    /// This is similar to the CameraSnapAndRender, however it doesn't use the rendering pipeline at all.
    /// </summary>
    public class CameraSnap : IDisposable
    {
        // ====================================================================

        #region Private members

        /// <summary>
        /// Private field. Use the public property <see cref="Moniker"/> for access to this value.
        /// </summary>
        private IMoniker _Moniker = null;

        /// <summary>
        /// Private field. Use the public property <see cref="Resolution"/> for access to this value.
        /// </summary>
        private Resolution _Resolution = null;

        /// <summary>
        /// Private field. Use the public property <see cref="ResolutionList"/> for access to this value.
        /// </summary>
        private ResolutionList _ResolutionList = new ResolutionList();


        /// <summary>
        /// Private field. Use the public property <see cref="VideoInput"/> for access to this value.
        /// </summary>
        private VideoInput _VideoInput = VideoInput.Default;

        /// <summary>
        /// Private field. Use the public property <see cref="DirectShowLogFilepath"/> for access to this value.
        /// </summary>
        private string _DirectShowLogFilepath = string.Empty;

        /// <summary>
        /// Private field for DirectShow log file handle.
        /// </summary>
        private IntPtr _DirectShowLogHandle = IntPtr.Zero;

        // Header's size of DIB image returned by IVMRWindowlessControl9::GetCurrentImage method (BITMAPINFOHEADER)
        private static readonly int DIB_Image_HeaderSize = Marshal.SizeOf(typeof(BitmapInfoHeader)); // == 40;

        #endregion

        // ====================================================================

        #region Internal stuff

        /// <summary>
        /// Private field. DirectXInterfaces instance.
        /// </summary>
        internal DirectXInterfaces DX = new DirectXInterfaces();


        /// <summary>
        /// Private field. Was the graph built or not.
        /// </summary>
        internal bool _bGraphIsBuilt = false;

        /// <summary>
        /// Private field. SampleGrabber helper (wrapper)
        /// </summary>
        internal SampleGrabberHelper _pSampleGrabberHelper = null;

        #if DEBUG
        /// <summary>
        /// Private field. DsROTEntry allows to "Connect to remote graph" from GraphEdit
        /// </summary>
        DsROTEntry _rot = null;
        #endif


        #endregion


        // ====================================================================

        #region Public properties

        /// <summary>
        /// Gets a camera moniker (device identification).
        /// </summary> 
        public IMoniker Moniker
        {
            get { return _Moniker; }
        }

        /// <summary>
        /// Gets or sets a resolution of camera's output.
        /// </summary>
        /// <seealso cref="ResolutionListRGB"/>
        public Resolution Resolution
        {
            get { return _Resolution; }
            set
            {
                // Change of resolution is not allowed after graph's built
                if (_bGraphIsBuilt)
                    throw new Exception(@"Change of resolution is not allowed after graph's built.");

                _Resolution = value;
            }
        }

        /// <summary>
        /// Gets a list of available resolutions (in RGB format).
        /// </summary>        
        public ResolutionList ResolutionListRGB
        {
            get { return _ResolutionList; }
        }


        /// <summary>
        /// Log file path for directshow (used in BuildGraph)
        /// </summary> 
        /// <seealso cref="BuildGraph"/>
        public string DirectShowLogFilepath
        {
            get
            {
                return _DirectShowLogFilepath;
            }
            set
            {
                _DirectShowLogFilepath = value;

                ApplyDirectShowLogFile();
            }
        }
        #endregion


        // ====================================================================

        #region Public member functions

        #region Create, Initialize and Dispose/Close

        /// <summary>
        /// Default constructor for <see cref="CameraSnap"/> class.
        /// </summary>
        public CameraSnap()
        {
        }

        /// <summary>
        /// Initializes camera and connects it to HostingControl and Moniker.
        /// </summary>
        /// <param name="hControl">Control that is used for hosting camera's output.</param>
        /// <param name="moniker">Moniker (device identification) of camera.</param>
        /// <seealso cref="HostingControl"/>
        /// <seealso cref="Moniker"/>
        public void Initialize(IMoniker moniker)
        {
            if ( moniker == null )
                throw new Exception(@"Camera's moniker should be set.");

            _Moniker = moniker;
        }

        /// <summary>
        /// Destructor (disposer) for <see cref="CameraSnap"/> class.
        /// </summary>
        ~CameraSnap()
        {
            Dispose();
        }

        /// <summary>
        /// Dispose of <see cref="IDisposable"/> for <see cref="CameraSnap"/> class.
        /// </summary>
        public void Dispose()
        {
            CloseAll();
        }

        /// <summary>
        /// Close and dispose all camera and DirectX stuff.
        /// </summary>
        public void CloseAll()
        {
            _bGraphIsBuilt = false;

            // close log file if needed
            try
            {
                CloseDirectShowLogFile();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            // stop rendering
            if (DX.MediaControl != null)
            {
                try
                {
                    DX.MediaControl.StopWhenReady();
                    DX.MediaControl.Stop();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }
            }

            //FilterGraphTools.RemoveAllFilters(this.graphBuilder);

#if DEBUG
            if (_rot != null)
            {
                _rot.Dispose();
            }
#endif

            // Dispose Managed Direct3D objects
            if (_pSampleGrabberHelper != null)
            {
                _pSampleGrabberHelper.Dispose();
                _pSampleGrabberHelper = null;
            }

            DX.CloseInterfaces();

        }

        #endregion

        #region Graph: Build, Run, Stop



        /// <summary>
        /// Builds DirectShow graph for rendering.
        /// </summary>
        public void BuildGraph(Action<Bitmap> act, RotateFlipType rft)
        {
            _bGraphIsBuilt = false;

            try
            {
                // -------------------------------------------------------
                DX.FilterGraph = (IFilterGraph2)new FilterGraph();
                DX.MediaControl = (IMediaControl)DX.FilterGraph;

                // Log file if needed
                ApplyDirectShowLogFile();

#if DEBUG
                // Allows you to view the graph with GraphEdit File/Connect
                _rot = new DsROTEntry(DX.FilterGraph);
#endif

                // -------------------------------------------------------
                GraphBuilding_AddFilter_Source();
                GraphBuilding_SetSourceParams();
                GraphBuilding_AddFilter_SampleGrabber(act, rft);
                GraphBuilding_ConnectPins();

                // -------------------------------------------------------
                this._Resolution = _pSampleGrabberHelper.SaveMode();

                // -------------------------------------------------------

                // -------------------------------------------------------
                _bGraphIsBuilt = true;
                // -------------------------------------------------------

            }
            catch
            {
                CloseAll();
                throw;
            }

#if DEBUG
            // Double check to make sure we aren't releasing something
            // important.
            GC.Collect();
            GC.WaitForPendingFinalizers();
#endif
        }

        /// <summary>
        /// Runs DirectShow graph for rendering.
        /// </summary>
        public void RunGraph()
        {
            //var graph_guilder = (ICaptureGraphBuilder2)new CaptureGraphBuilder2();
            //int hr = graph_guilder.RenderStream(PinCategory.Preview, MediaType.Video, DX.CaptureFilter, null, DX.VMRenderer);
            //DsError.ThrowExceptionForHR(hr);

            if (DX.MediaControl != null)
            {
                int hr = DX.MediaControl.Run();
                DsError.ThrowExceptionForHR(hr);
            }

        }

        /// <summary>
        /// Stops DirectShow graph for rendering.
        /// </summary>
        public void StopGraph()
        {
            if (DX.MediaControl != null)
            {
                int hr = DX.MediaControl.Stop();
                DsError.ThrowExceptionForHR(hr);
            }
        }
        #endregion

        #region Property pages (various settings dialogs)


        /// <summary>
        /// Displays property page for capture filter.
        /// </summary>
        /// <param name="hwndOwner">The window handler for to make it parent of property page.</param>
        public void DisplayPropertyPage_CaptureFilter(IntPtr hwndOwner)
        {
            CameraHelpers.DisplayPropertyPageFilter(DX.CaptureFilter, hwndOwner);
        }

        /// <summary>
        /// Displays property page for filter's pin output.
        /// </summary>
        /// <param name="hwndOwner">The window handler for to make it parent of property page.</param>
        public void DisplayPropertyPage_SourcePinOutput(IntPtr hwndOwner)
        {
            IPin pinSourceCapture = null;

            try
            {
                pinSourceCapture = DsFindPin.ByDirection(DX.CaptureFilter, PinDirection.Output, 0);
                CameraHelpers.DisplayPropertyPagePin(pinSourceCapture, hwndOwner);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                CameraHelpers.SafeReleaseComObject(pinSourceCapture);
                pinSourceCapture = null;
            }
        }


        #endregion

        // ====================================================================

        #region Graph building stuff


        /// <summary>
        /// Sets the Framerate, and video size.
        /// </summary>
        private void GraphBuilding_SetSourceParams()
        {
            // Pins used in graph
            IPin pinSourceCapture = null;

            try
            {
                // Collect pins
                //pinSourceCapture = DsFindPin.ByCategory(DX.CaptureFilter, PinCategory.Capture, 0);
                pinSourceCapture = DsFindPin.ByDirection(DX.CaptureFilter, PinDirection.Output, 0);

                CameraHelpers.SetSourceParams(pinSourceCapture, _Resolution);
            }
            catch
            {
                throw;
            }
            finally
            {
                CameraHelpers.SafeReleaseComObject(pinSourceCapture);
                pinSourceCapture = null;
            }
        }

        /// <summary>
        /// Connects pins of graph
        /// </summary>
        private void GraphBuilding_ConnectPins()
        {
            // Pins used in graph
            IPin pinSourceCapture = null;

            IPin pinSampleGrabberInput = null;

            int hr = 0;

            try
            {
                // Collect pins
                //pinSourceCapture = DsFindPin.ByCategory(DX.CaptureFilter, PinCategory.Capture, 0);
                pinSourceCapture = DsFindPin.ByDirection(DX.CaptureFilter, PinDirection.Output, 0);

                pinSampleGrabberInput = DsFindPin.ByDirection(DX.SampleGrabberFilter, PinDirection.Input, 0);

                // Connect source to tee splitter
                hr = DX.FilterGraph.Connect(pinSourceCapture, pinSampleGrabberInput);
                DsError.ThrowExceptionForHR(hr);
            }
            catch
            {
                throw;
            }
            finally
            {
                CameraHelpers.SafeReleaseComObject(pinSourceCapture);
                pinSourceCapture = null;

                CameraHelpers.SafeReleaseComObject(pinSampleGrabberInput);
                pinSampleGrabberInput = null;
            }
        }
        
        #endregion

        #region Filters

        /// <summary>
        /// Adds video source filter to the filter graph.
        /// </summary>
        private void GraphBuilding_AddFilter_Source()
        {
            int hr = 0;

            DX.CaptureFilter = null;
            hr = DX.FilterGraph.AddSourceFilterForMoniker(_Moniker, null, "Source Filter", out DX.CaptureFilter);
            DsError.ThrowExceptionForHR(hr);

            _ResolutionList = CameraHelpers.GetResolutionsAvailable(DX.CaptureFilter);
        }
                
        /// <summary>
        /// Adds SampleGrabber for screenshot making.
        /// </summary>
        private void GraphBuilding_AddFilter_SampleGrabber(Action<Bitmap> bp, RotateFlipType rft)
        {
            int hr = 0;

            // Get the SampleGrabber interface
            DX.SampleGrabber = new SampleGrabber() as ISampleGrabber;
            
            // Configure the sample grabber
            DX.SampleGrabberFilter = DX.SampleGrabber as IBaseFilter;
            _pSampleGrabberHelper = new SampleGrabberHelper(DX.SampleGrabber, bp, rft);

            _pSampleGrabberHelper.ConfigureMode();

            // Add the sample grabber to the graph
            hr = DX.FilterGraph.AddFilter(DX.SampleGrabberFilter, "Sample Grabber");
            DsError.ThrowExceptionForHR(hr);
        }


        #endregion

        #region Internal event handlers for HostingControl and system


        /// <summary>
        /// Handler of SystemEvents.DisplaySettingsChanged.
        /// </summary>
        private void SystemEvents_DisplaySettingsChanged(object sender, EventArgs e)
        {
            if (!_bGraphIsBuilt)
                return; // Do nothing before graph was built

            if (DX.WindowlessCtrl != null)
            {
                int hr = DX.WindowlessCtrl.DisplayModeChanged();
            }
        }

        #endregion


        

        #region Private Helpers

        /// <summary>
        /// Open log file for DirectShow.
        /// </summary>
        private void ApplyDirectShowLogFile()
        {
            if (DX.FilterGraph == null)
                return; // can't be set now. Will be set on BuildGraph()

            CloseDirectShowLogFile();

            if (string.IsNullOrEmpty(_DirectShowLogFilepath))
            {
                return;
            }

            _DirectShowLogHandle = NativeMethods.CreateFile(_DirectShowLogFilepath,
                System.IO.FileAccess.Write,
                System.IO.FileShare.Read,
                IntPtr.Zero,
                System.IO.FileMode.OpenOrCreate,
                System.IO.FileAttributes.Normal,
                IntPtr.Zero);


            if (_DirectShowLogHandle.ToInt32() == -1 || _DirectShowLogHandle == IntPtr.Zero)
            {
                _DirectShowLogHandle = IntPtr.Zero;

                throw new Exception("Can't open log file for writing: " + _DirectShowLogFilepath);
            }

            // Append to file - move to end (WinApi's CreateFile doesn't support append FileMode)
            NativeMethods.SetFilePointerEx(_DirectShowLogHandle, 0, IntPtr.Zero, NativeMethods.FILE_END);


            DX.FilterGraph.SetLogFile(_DirectShowLogHandle);
        }

        /// <summary>
        /// Close log file for DirectShow.
        /// </summary>
        private void CloseDirectShowLogFile()
        {
            try
            {
                if (DX.FilterGraph != null)
                {
                    DX.FilterGraph.SetLogFile(IntPtr.Zero);
                }

                NativeMethods.CloseHandle(_DirectShowLogHandle);
            }
            catch
            {
                throw;
            }
            finally
            {
                _DirectShowLogHandle = IntPtr.Zero;
            }

        }
        
        #endregion
        

        #endregion // Private

        // ====================================================================
    }
}