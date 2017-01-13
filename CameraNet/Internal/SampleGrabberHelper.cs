#region License

/*
CameraNet - Camera wrapper for directshow for .NET
Copyright (C) 2013
https://github.com/free5lot/CameraNet

This library is free software; you can redistribute it and/or
modify it under the terms of the GNU Lesser General Public
License as published by the Free Software Foundation; either
version 3.0 of the License, or (at your option) any later version.

This library is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
Lesser General Public License for more details.

You should have received a copy of the GNU LesserGeneral Public 
License along with this library. If not, see <http://www.gnu.org/licenses/>.
*/

#endregion

namespace CameraNet
{
    #region Using directives

    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Runtime.InteropServices;
    using System.Diagnostics;
    using System.Threading;
    using System.Drawing;
    using System.Drawing.Imaging;

    // Use DirectShowLib (LGPL v2.1)
    using DirectShowLib;

    #endregion

    /// <summary>
    /// Helper for SampleGrabber. Used to make screenshots (snapshots).
    /// </summary>
    /// <remarks>This class is inherited from <see cref="ISampleGrabberCB"/> class.</remarks>
    /// 
    /// <author> free5lot (free5lot@yandex.ru) </author>
    /// <version> 2013.10.17 </version>
    internal sealed class SampleGrabberHelper : ISampleGrabberCB, IDisposable
    {
        private RotateFlipType _RotateFlipType;
        private Action<Bitmap> _RawImageEventHandler = null;

        private int _ResolutionWidth;
        private int _ResolutionHeight;
        private int _ResolutionBitsPerPixel;
        private int _ResolutionSizeBytes;

        /// <summary>
        /// Pointer to COM-interface ISampleGrabber.
        /// </summary>
        private ISampleGrabber _SampleGrabber = null;
        private readonly object _DeltaLock = new object();

        /// <summary>
        /// Default constructor for <see cref="SampleGrabberHelper"/> class.
        /// </summary>
        /// <param name="sampleGrabber">Pointer to COM-interface ISampleGrabber.</param>
        public SampleGrabberHelper(ISampleGrabber sampleGrabber, Action<Bitmap> raw_image_handler, RotateFlipType rft = RotateFlipType.RotateNoneFlipNone)
        {
            _SampleGrabber = sampleGrabber;
            this._RawImageEventHandler = raw_image_handler;
            this._RotateFlipType = rft;
        }

        /// <summary>
        /// Disposes object and snapshot.
        /// </summary>
        public void Dispose()
        {
            lock (_DeltaLock)
            {
                this._RawImageEventHandler = null;
                this._SampleGrabber = null;
            }
        }

        /// <summary>
        /// Configures mode (mediatype, format type and etc).
        /// </summary>
        public void ConfigureMode()
        {
            lock (_DeltaLock)
            {
                int hr;
                AMMediaType media = new AMMediaType();

                // Set the media type to Video/RBG24
                media.majorType = MediaType.Video;
                media.subType = MediaSubType.RGB24;
                media.formatType = FormatType.VideoInfo;
                hr = _SampleGrabber.SetMediaType(media);
                DsError.ThrowExceptionForHR(hr);

                DsUtils.FreeAMMediaType(media);
                media = null;

                // Configure the samplegrabber

                // To save current frame via SnapshotNextFrame
                //ISampleGrabber::SetCallback method
                // Note  [Deprecated. This API may be removed from future releases of Windows.]
                // http://msdn.microsoft.com/en-us/library/windows/desktop/dd376992%28v=vs.85%29.aspx
                hr = _SampleGrabber.SetCallback(this, 1); // 1 == WhichMethodToCallback, call the ISampleGrabberCB::BufferCB method
                DsError.ThrowExceptionForHR(hr);
            }
        }

        /// <summary>
        /// Gets and saves mode (mediatype, format type and etc). 
        /// </summary>
        public Resolution SaveMode()
        {
            lock (_DeltaLock)
            {
                int hr;

                // Get the media type from the SampleGrabber
                AMMediaType media = new AMMediaType();

                hr = _SampleGrabber.GetConnectedMediaType(media);
                DsError.ThrowExceptionForHR(hr);

                if ((media.formatType != FormatType.VideoInfo) || (media.formatPtr == IntPtr.Zero))
                {
                    throw new NotSupportedException("Unknown Grabber Media Format");
                }

                // Grab the size info
                VideoInfoHeader videoInfoHeader = (VideoInfoHeader)Marshal.PtrToStructure(media.formatPtr, typeof(VideoInfoHeader));
                _ResolutionWidth = videoInfoHeader.BmiHeader.Width;
                _ResolutionHeight = videoInfoHeader.BmiHeader.Height;
                _ResolutionBitsPerPixel = videoInfoHeader.BmiHeader.BitCount;
                _ResolutionSizeBytes = videoInfoHeader.BmiHeader.ImageSize;

                DsUtils.FreeAMMediaType(media);
                media = null;

                return new Resolution(_ResolutionWidth, _ResolutionHeight);
            }
        }

        /// <summary>
        /// SampleCB callback (NOT USED). It should be implemented for ISampleGrabberCB
        /// </summary>
        int ISampleGrabberCB.SampleCB(double SampleTime, IMediaSample pSample)
        {
            Marshal.ReleaseComObject(pSample);
            return 0;
        }



        /// <summary>
        /// BufferCB callback 
        /// </summary>
        /// <remarks>COULD BE EXECUTED FROM FOREIGN THREAD.</remarks>
        int ISampleGrabberCB.BufferCB(double SampleTime, IntPtr pBuffer, int BufferLen)
        {
            Bitmap bitmap_clone = null;

            lock (_DeltaLock)
            {
                if (_RawImageEventHandler == null)
                {
                    return 0;
                }
                if (BufferLen == 0)
                {
                    Console.WriteLine(this.GetType().FullName.ToString() + "BufferCB(...) Buffer Length 0");
                    return 0;
                }
                IntPtr m_ipBuffer = IntPtr.Zero;

                // get ready to wait for new image
                try
                {
                    m_ipBuffer = Marshal.AllocCoTaskMem(Math.Abs(_ResolutionBitsPerPixel / 8 * _ResolutionWidth) * _ResolutionHeight);
                    NativeMethods.CopyMemory(m_ipBuffer, pBuffer, BufferLen);

                    PixelFormat pixelFormat = PixelFormat.Format24bppRgb;
                    switch (_ResolutionBitsPerPixel)
                    {
                        case 24:
                            pixelFormat = PixelFormat.Format24bppRgb;
                            break;
                        case 32:
                            pixelFormat = PixelFormat.Format32bppRgb;
                            break;
                        case 48:
                            pixelFormat = PixelFormat.Format48bppRgb;
                            break;
                        default:
                            throw new Exception("Unsupported BitCount");
                    }

                    Bitmap bitmap = new Bitmap(_ResolutionWidth, _ResolutionHeight, (_ResolutionBitsPerPixel / 8) * _ResolutionWidth, pixelFormat, m_ipBuffer);

                    bitmap_clone = bitmap.Clone(new Rectangle(0, 0, _ResolutionWidth, _ResolutionHeight), PixelFormat.Format24bppRgb);
                    bitmap_clone.RotateFlip(this._RotateFlipType);

                    Marshal.FreeCoTaskMem(m_ipBuffer);

                    bitmap.Dispose();
                    bitmap = null;
                }
                catch
                {
                    Marshal.FreeCoTaskMem(m_ipBuffer);
                }

            }

            if (bitmap_clone != null)
                this._RawImageEventHandler(bitmap_clone);
            return 0;
        }

        internal void UpdateRotateFlipType(RotateFlipType rft)
        {
            lock (_DeltaLock)
                this._RotateFlipType = rft;
        }
    }
}
