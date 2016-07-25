using DirectShowLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace CameraNet
{
    public static class CameraHelpers
    {

        /// <summary>
        /// public. Displays a property page for a filter
        /// </summary>
        /// <param name="filter">The filter for which to display a property page.</param>
        /// <param name="hwndOwner">The window handler for to make it parent of property page.</param>
        public static void DisplayPropertyPageFilter(IBaseFilter filter, IntPtr hwndOwner)
        {
            _DisplayPropertyPage(filter, hwndOwner);
        }
        public static void DisplayPropertyPagePin(IPin pin, IntPtr hwndOwner)
        {
            _DisplayPropertyPage(pin, hwndOwner);
        }

        public static void _DisplayPropertyPage(object filter_or_pin, IntPtr hwndOwner)
        {
            if (filter_or_pin == null)
                return;

            //Get the ISpecifyPropertyPages for the filter
            ISpecifyPropertyPages pProp = filter_or_pin as ISpecifyPropertyPages;
            int hr = 0;

            if (pProp == null)
            {
                //If the filter doesn't implement ISpecifyPropertyPages, try displaying IAMVfwCompressDialogs instead!
                IAMVfwCompressDialogs compressDialog = filter_or_pin as IAMVfwCompressDialogs;
                if (compressDialog != null)
                {

                    hr = compressDialog.ShowDialog(VfwCompressDialogs.Config, IntPtr.Zero);
                    DsError.ThrowExceptionForHR(hr);
                }
                return;
            }

            string caption = string.Empty;

            if (filter_or_pin is IBaseFilter)
            {
                //Get the name of the filter from the FilterInfo struct
                IBaseFilter as_filter = filter_or_pin as IBaseFilter;
                FilterInfo filterInfo;
                hr = as_filter.QueryFilterInfo(out filterInfo);
                DsError.ThrowExceptionForHR(hr);

                caption = filterInfo.achName;

                if (filterInfo.pGraph != null)
                {
                    Marshal.ReleaseComObject(filterInfo.pGraph);
                }
            }
            else
            if (filter_or_pin is IPin)
            {
                //Get the name of the filter from the FilterInfo struct
                IPin as_pin = filter_or_pin as IPin;
                PinInfo pinInfo;
                hr = as_pin.QueryPinInfo(out pinInfo);
                DsError.ThrowExceptionForHR(hr);

                caption = pinInfo.name;
            }


            // Get the propertypages from the property bag
            DsCAUUID caGUID;
            hr = pProp.GetPages(out caGUID);
            DsError.ThrowExceptionForHR(hr);

            // Create and display the OlePropertyFrame
            object oDevice = (object)filter_or_pin;
            hr = NativeMethods.OleCreatePropertyFrame(hwndOwner, 0, 0, caption, 1, ref oDevice, caGUID.cElems, caGUID.pElems, 0, 0, IntPtr.Zero);
            DsError.ThrowExceptionForHR(hr);

            // Release COM objects
            Marshal.FreeCoTaskMem(caGUID.pElems);
            Marshal.ReleaseComObject(pProp);
        }


        // ====================================================================

        #region public Static functions

        /// <summary>
        /// Returns Moniker (device identification) of camera from device index.
        /// </summary>
        /// <param name="iDeviceIndex">Index (Zero-based) in list of available devices with VideoInputDevice filter category.</param>
        /// <returns>Moniker (device identification) of device</returns>
        public static IMoniker GetDeviceMoniker(int iDeviceIndex)
        {
            DsDevice[] capDevices;

            // Get the collection of video devices
            capDevices = DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice);

            if (iDeviceIndex >= capDevices.Length)
            {
                throw new Exception(@"No video capture devices found at that index.");
            }

            return capDevices[iDeviceIndex].Mon;
        }

        /// <summary>
        /// Returns available resolutions with RGB color system for device moniker
        /// </summary>
        /// <param name="moniker">Moniker (device identification) of camera.</param>
        /// <returns>List of resolutions with RGB color system of device</returns>
        public static ResolutionList GetResolutionList(IMoniker moniker)
        {
            int hr;

            ResolutionList ResolutionsAvailable = null; //new ResolutionList();

            // Get the graphbuilder object
            IFilterGraph2 filterGraph = new FilterGraph() as IFilterGraph2;
            IBaseFilter capFilter = null;

            try
            {
                // add the video input device
                hr = filterGraph.AddSourceFilterForMoniker(moniker, null, "Source Filter", out capFilter);
                DsError.ThrowExceptionForHR(hr);

                ResolutionsAvailable = GetResolutionsAvailable(capFilter);
            }
            finally
            {
                SafeReleaseComObject(filterGraph);
                filterGraph = null;

                SafeReleaseComObject(capFilter);
                capFilter = null;
            }

            return ResolutionsAvailable;
        }

        #endregion


        #region Resolution lists

        /// <summary>
        /// Gets available resolutions (which are appropriate for us) for capture filter.
        /// </summary>
        /// <param name="captureFilter">Capture filter for asking for resolution list.</param>
        public static ResolutionList GetResolutionsAvailable(IBaseFilter captureFilter)
        {
            ResolutionList resolution_list = null;

            IPin pRaw = null;
            try
            {
                pRaw = DsFindPin.ByDirection(captureFilter, PinDirection.Output, 0);
                //pRaw = DsFindPin.ByCategory(captureFilter, PinCategory.Capture, 0);
                //pRaw = DsFindPin.ByCategory(filter, PinCategory.Preview, 0);

                resolution_list = GetResolutionsAvailable(pRaw);
            }
            catch
            {
                throw;
                //resolution_list = new ResolutionList();
                //resolution_list.Add(new Resolution(640, 480));
            }
            finally
            {
                SafeReleaseComObject(pRaw);
                pRaw = null;
            }

            return resolution_list;
        }

        /// <summary>
        /// Free media type if needed.
        /// </summary>
        /// <param name="media_type">Media type to free.</param>
        public static void FreeMediaType(ref AMMediaType media_type)
        {
            if (media_type == null)
                return;

            DsUtils.FreeAMMediaType(media_type);
            media_type = null;
        }

        /// <summary>
        /// Free SCC (it's not used but required for GetStreamCaps()).
        /// </summary>
        /// <param name="pSCC">SCC to free.</param>
        public static void FreeSCCMemory(ref IntPtr pSCC)
        {
            if (pSCC == IntPtr.Zero)
                return;

            Marshal.FreeCoTaskMem(pSCC);
            pSCC = IntPtr.Zero;
        }


        /// <summary>
        /// Gets available resolutions (which are appropriate for us) for capture pin (PinCategory.Capture).
        /// </summary>
        /// <param name="captureFilter">Capture pin (PinCategory.Capture) for asking for resolution list.</param>
        public static ResolutionList GetResolutionsAvailable(IPin pinOutput)
        {
            int hr = 0;

            ResolutionList ResolutionsAvailable = new ResolutionList();

            //ResolutionsAvailable.Clear();

            // Media type (shoudl be cleaned)
            AMMediaType media_type = null;

            //NOTE: pSCC is not used. All we need is media_type
            IntPtr pSCC = IntPtr.Zero;

            try
            {
                IAMStreamConfig videoStreamConfig = pinOutput as IAMStreamConfig;

                // -------------------------------------------------------------------------
                // We want the interface to expose all media types it supports and not only the last one set
                hr = videoStreamConfig.SetFormat(null);
                DsError.ThrowExceptionForHR(hr);

                int piCount = 0;
                int piSize = 0;

                hr = videoStreamConfig.GetNumberOfCapabilities(out piCount, out piSize);
                DsError.ThrowExceptionForHR(hr);

                for (int i = 0; i < piCount; i++)
                {
                    // ---------------------------------------------------
                    pSCC = Marshal.AllocCoTaskMem(piSize);
                    videoStreamConfig.GetStreamCaps(i, out media_type, pSCC);

                    // NOTE: we could use VideoStreamConfigCaps.InputSize or something like that to get resolution, but it's deprecated
                    //VideoStreamConfigCaps videoStreamConfigCaps = (VideoStreamConfigCaps)Marshal.PtrToStructure(pSCC, typeof(VideoStreamConfigCaps));
                    // ---------------------------------------------------

                    if (IsBitCountAppropriate(GetBitCountForMediaType(media_type)))
                    {
                        ResolutionsAvailable.AddIfNew(GetResolutionForMediaType(media_type));
                    }

                    FreeSCCMemory(ref pSCC);
                    FreeMediaType(ref media_type);
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                // clean up
                FreeSCCMemory(ref pSCC);
                FreeMediaType(ref media_type);
            }

            return ResolutionsAvailable;
        }

        #endregion

        /// <summary>
        /// Displays property page for device.
        /// </summary>
        /// <param name="moniker">Moniker (device identification) of camera.</param>
        /// <param name="hwndOwner">The window handler for to make it parent of property page.</param>
        /// <seealso cref="Moniker"/>
        public static void DisplayPropertyPage_Device(IMoniker moniker, IntPtr hwndOwner)
        {
            if (moniker == null)
                return;

            object source = null;
            Guid iid = typeof(IBaseFilter).GUID;
            moniker.BindToObject(null, null, ref iid, out source);
            IBaseFilter theDevice = (IBaseFilter)source;

            DisplayPropertyPageFilter(theDevice, hwndOwner);

            //Release COM objects
            SafeReleaseComObject(theDevice);
            theDevice = null;
        }


        /// <summary>
        /// Checks if AMMediaType's resolution is appropriate for desired resolution.
        /// </summary>
        /// <param name="media_type">Media type to analyze.</param>
        /// <param name="resolution_desired">Desired resolution. Can be null or have 0 for height or width if it's not important.</param>
        public static bool IsResolutionAppropiate(AMMediaType media_type, Resolution resolution_desired)
        {
            // if we were asked to choose resolution
            if (resolution_desired == null)
                return true;

            VideoInfoHeader videoInfoHeader = new VideoInfoHeader();
            Marshal.PtrToStructure(media_type.formatPtr, videoInfoHeader);

            if (resolution_desired.Width > 0 &&
                videoInfoHeader.BmiHeader.Width != resolution_desired.Width)
            {
                return false;
            }
            if (resolution_desired.Height > 0 &&
                videoInfoHeader.BmiHeader.Height != resolution_desired.Height)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Get resoltuin from if AMMediaType's resolution is appropriate for resolution_desired
        /// </summary>
        /// <param name="media_type">Media type to analyze.</param>
        /// <param name="resolution_desired">Desired resolution. Can be null or have 0 for height or width if it's not important.</param>
        public static Resolution GetResolutionForMediaType(AMMediaType media_type)
        {
            VideoInfoHeader videoInfoHeader = new VideoInfoHeader();
            Marshal.PtrToStructure(media_type.formatPtr, videoInfoHeader);

            return new Resolution(videoInfoHeader.BmiHeader.Width, videoInfoHeader.BmiHeader.Height);
        }


        /// <summary>
        /// Get bit count for mediatype
        /// </summary>
        /// <param name="media_type">Media type to analyze.</param>
        public static short GetBitCountForMediaType(AMMediaType media_type)
        {

            VideoInfoHeader videoInfoHeader = new VideoInfoHeader();
            Marshal.PtrToStructure(media_type.formatPtr, videoInfoHeader);

            return videoInfoHeader.BmiHeader.BitCount;
        }

        /// <summary>
        /// Check if bit count is appropriate for us
        /// </summary>
        /// <param name="media_type">Media type to analyze.</param>
        public static bool IsBitCountAppropriate(short bit_count)
        {
            if (bit_count == 16 ||
                bit_count == 24 ||
                bit_count == 32)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Analyze AMMediaType during enumeration and decide if it's good choice for us.
        /// </summary>
        /// <param name="media_type">Media type to analyze.</param>
        /// <param name="resolution_desired">Desired resolution.</param>
        public static void AnalyzeMediaType(AMMediaType media_type, Resolution resolution_desired, out bool bit_count_ok, out bool sub_type_ok, out bool resolution_ok)
        {
            // ---------------------------------------------------
            short bit_count = GetBitCountForMediaType(media_type);

            bit_count_ok = IsBitCountAppropriate(bit_count);

            // ---------------------------------------------------

            // We want (A)RGB32, RGB24 or RGB16 and YUY2.
            // These have priority
            // Change this if you're not agree.
            sub_type_ok = (
                media_type.subType == MediaSubType.RGB32 ||
                media_type.subType == MediaSubType.ARGB32 ||
                media_type.subType == MediaSubType.RGB24 ||
                media_type.subType == MediaSubType.RGB16_D3D_DX9_RT ||
                media_type.subType == MediaSubType.RGB16_D3D_DX7_RT ||
                media_type.subType == MediaSubType.YUY2);

            // ---------------------------------------------------

            // flag to show if media_type's resolution is appropriate for us
            resolution_ok = IsResolutionAppropiate(media_type, resolution_desired);
            // ---------------------------------------------------
        }

        /// <summary>
        /// Sets parameters for source capture pin.
        /// </summary>
        /// <param name="pinSourceCapture">Pin of source capture.</param>
        /// <param name="resolution">Resolution to set if possible.</param>
        public static void SetSourceParams(IPin pinSourceCapture, Resolution resolution_desired)
        {
            int hr = 0;

            AMMediaType media_type_most_appropriate = null;
            AMMediaType media_type = null;

            //NOTE: pSCC is not used. All we need is media_type
            IntPtr pSCC = IntPtr.Zero;


            bool appropriate_media_type_found = false;

            try
            {
                IAMStreamConfig videoStreamConfig = pinSourceCapture as IAMStreamConfig;

                // -------------------------------------------------------------------------
                // We want the interface to expose all media types it supports and not only the last one set
                hr = videoStreamConfig.SetFormat(null);
                DsError.ThrowExceptionForHR(hr);

                int piCount = 0;
                int piSize = 0;

                hr = videoStreamConfig.GetNumberOfCapabilities(out piCount, out piSize);
                DsError.ThrowExceptionForHR(hr);

                for (int i = 0; i < piCount; i++)
                {
                    // ---------------------------------------------------
                    pSCC = Marshal.AllocCoTaskMem(piSize);
                    videoStreamConfig.GetStreamCaps(i, out media_type, pSCC);
                    FreeSCCMemory(ref pSCC);

                    // NOTE: we could use VideoStreamConfigCaps.InputSize or something like that to get resolution, but it's deprecated
                    //VideoStreamConfigCaps videoStreamConfigCaps = (VideoStreamConfigCaps)Marshal.PtrToStructure(pSCC, typeof(VideoStreamConfigCaps));
                    // ---------------------------------------------------

                    bool bit_count_ok = false;
                    bool sub_type_ok = false;
                    bool resolution_ok = false;

                    AnalyzeMediaType(media_type, resolution_desired, out bit_count_ok, out sub_type_ok, out resolution_ok);

                    if (bit_count_ok && resolution_ok)
                    {
                        if (sub_type_ok)
                        {
                            hr = videoStreamConfig.SetFormat(media_type);
                            DsError.ThrowExceptionForHR(hr);

                            appropriate_media_type_found = true;
                            break; // stop search, we've found appropriate media type
                        }
                        else
                        {
                            // save as appropriate if no other found
                            if (media_type_most_appropriate == null)
                            {
                                media_type_most_appropriate = media_type;
                                media_type = null; // we don't want for free it, now it's media_type_most_appropriate's problem
                            }
                        }
                    }

                    FreeMediaType(ref media_type);
                }

                if (!appropriate_media_type_found)
                {
                    // Found nothing exactly as we were asked 

                    if (media_type_most_appropriate != null)
                    {
                        // set appropriate RGB format with different resolution
                        hr = videoStreamConfig.SetFormat(media_type_most_appropriate);
                        DsError.ThrowExceptionForHR(hr);
                    }
                    else
                    {
                        // throw. We didn't find exactly what we were asked to
                        throw new Exception("Camera doesn't support media type with requested resolution and bits per pixel.");
                        //DsError.ThrowExceptionForHR(DsResults.E_InvalidMediaType);
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                // clean up
                FreeMediaType(ref media_type);
                FreeMediaType(ref media_type_most_appropriate);
                FreeSCCMemory(ref pSCC);

            }
        }

        #region Crossbar helpers

        /// <summary>
        /// Gets type of input connected to video output of the crossbar.
        /// </summary>
        /// <param name="crossbar">The crossbar of device.</param>
        /// <returns>Video input of device</returns>
        /// <seealso cref="CrossbarAvailable"/>
        public static VideoInput GetCrossbarInput(IAMCrossbar crossbar)
        {
            VideoInput videoInput = VideoInput.Default;

            int inPinsCount, outPinsCount;

            // gen number of pins in the crossbar
            if (crossbar.get_PinCounts(out outPinsCount, out inPinsCount) == 0)
            {
                int videoOutputPinIndex = -1;
                int pinIndexRelated;
                PhysicalConnectorType type;

                // find index of the video output pin
                for (int i = 0; i < outPinsCount; i++)
                {
                    if (crossbar.get_CrossbarPinInfo(false, i, out pinIndexRelated, out type) != 0)
                        continue;

                    if (type == PhysicalConnectorType.Video_VideoDecoder)
                    {
                        videoOutputPinIndex = i;
                        break;
                    }
                }

                if (videoOutputPinIndex != -1)
                {
                    int videoInputPinIndex;

                    // get index of the input pin connected to the output
                    if (crossbar.get_IsRoutedTo(videoOutputPinIndex, out videoInputPinIndex) == 0)
                    {
                        PhysicalConnectorType inputType;

                        crossbar.get_CrossbarPinInfo(true, videoInputPinIndex, out pinIndexRelated, out inputType);

                        videoInput = new VideoInput(videoInputPinIndex, inputType);
                    }
                }
            }

            return videoInput;
        }

        /// <summary>
        /// Sets type of input connected to video output of the crossbar.
        /// </summary>
        /// <param name="crossbar">The crossbar of device.</param>
        /// <param name="videoInput">Video input of device.</param>
        /// <seealso cref="CrossbarAvailable"/>
        public static void SetCrossbarInput(IAMCrossbar crossbar, VideoInput videoInput)
        {
            if (videoInput.Type != VideoInput.PhysicalConnectorType_Default &&
                videoInput.Index != -1)
            {
                int inPinsCount, outPinsCount;

                // gen number of pins in the crossbar
                if (crossbar.get_PinCounts(out outPinsCount, out inPinsCount) == 0)
                {
                    int videoOutputPinIndex = -1;
                    int videoInputPinIndex = -1;
                    int pinIndexRelated;
                    PhysicalConnectorType type;

                    // find index of the video output pin
                    for (int i = 0; i < outPinsCount; i++)
                    {
                        if (crossbar.get_CrossbarPinInfo(false, i, out pinIndexRelated, out type) != 0)
                            continue;

                        if (type == PhysicalConnectorType.Video_VideoDecoder)
                        {
                            videoOutputPinIndex = i;
                            break;
                        }
                    }

                    // find index of the required input pin
                    for (int i = 0; i < inPinsCount; i++)
                    {
                        if (crossbar.get_CrossbarPinInfo(true, i, out pinIndexRelated, out type) != 0)
                            continue;

                        if ((type == videoInput.Type) && (i == videoInput.Index))
                        {
                            videoInputPinIndex = i;
                            break;
                        }
                    }

                    // try connecting pins
                    if ((videoInputPinIndex != -1) && (videoOutputPinIndex != -1))
                    {
                        if (crossbar.CanRoute(videoOutputPinIndex, videoInputPinIndex) == 0)
                        {
                            int hr = crossbar.Route(videoOutputPinIndex, videoInputPinIndex);
                            DsError.ThrowExceptionForHR(hr);
                        }
                        else
                        {
                            throw new Exception("Can't route from selected VideoInput to VideoDecoder.");
                        }
                    }
                    else
                    {
                        throw new Exception("Can't find routing pins.");
                    }

                }
            }
        }

        #endregion

        /// <summary>
        /// Releases COM object
        /// </summary>
        /// <param name="obj">COM object to release.</param>
        public static void SafeReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
            }
        }
    }
}
