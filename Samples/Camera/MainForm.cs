using System;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices.ComTypes;


using CameraNet;

namespace SampleProject
{
    public partial class MainForm : Form
    {
        // Camera choice
        private CameraChoice _CameraChoice = new CameraChoice();
        private CameraSnap _Camera;

        public MainForm()
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            this.CloseCamera();
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        public void CloseCamera()
        {
            if (_Camera != null)
            {
                _Camera.StopGraph();
                _Camera.CloseAll();
                _Camera.Dispose();
                _Camera = null;
            }
        }

        // On load of Form
        private void FormCameraControlTool_Load(object sender, EventArgs e)
        {
            // Fill camera list combobox with available cameras
            FillCameraList();

            // Select the first one
            if (comboBoxCameraList.Items.Count > 0)
            {
                comboBoxCameraList.SelectedIndex = 0;
            }

            // Fill camera list combobox with available resolutions
            FillResolutionList();

            this.Enabled = true;

            this.timer1.Enabled = true;
        }

        // On close of Form
        private void FormMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.CloseCamera();
        }

        // Set current camera to camera_device
        private void SetCamera(IMoniker moniker, Resolution resolution = null)
        {
            try
            {
                // Makes all magic with camera and DirectShow graph
                // Close current if it was opened
                CloseCamera();

                if (moniker == null)
                    return;

                // Create camera object
                _Camera = new CameraSnap();

                string _DirectShowLogFilepath = string.Empty;

                if (!string.IsNullOrEmpty(_DirectShowLogFilepath))
                    _Camera.DirectShowLogFilepath = _DirectShowLogFilepath;

                // select resolution
                //ResolutionList resolutions = Camera.GetResolutionList(moniker);

                if (resolution != null)
                {
                    _Camera.Resolution = resolution;
                }

                // Initialize
                _Camera.Initialize(moniker);

                // Build and Run graph
                _Camera.BuildGraph();
                _Camera.RunGraph();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, @"Error while running camera");
            }
        }

        private void buttonCameraSettings_Click(object sender, EventArgs e)
        {
            if (this._Camera == null) return;

            CameraSnap.DisplayPropertyPage_Device(this._Camera.Moniker, this.Handle);
        }

        private void comboBoxCameraList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxCameraList.SelectedIndex < 0)
            {
                this.CloseCamera();
            }
            else
            {
                // Set camera
                this.SetCamera(_CameraChoice.Devices[comboBoxCameraList.SelectedIndex].Mon, null);
            }

            FillResolutionList();
        }

        private void comboBoxResolutionList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this._Camera == null) return;

            int comboBoxResolutionIndex = comboBoxResolutionList.SelectedIndex;
            if (comboBoxResolutionIndex < 0)
            {
                return;
            }
            ResolutionList resolutions = CameraSnap.GetResolutionList(this._Camera.Moniker);

            if (resolutions == null)
                return;

            if (comboBoxResolutionIndex >= resolutions.Count)
                return; // throw

            if (0 == resolutions[comboBoxResolutionIndex].CompareTo(this._Camera.Resolution))
            {
                // this resolution is already selected
                return;
            }

            // Recreate camera
            this.SetCamera(this._Camera.Moniker, resolutions[comboBoxResolutionIndex]);
        }

        private void FillResolutionList()
        {
            comboBoxResolutionList.Items.Clear();
            this.comboBoxResolutionList.Enabled = false;

            if (this._Camera == null) return;
            if (this._Camera.Resolution == null) return;

            ResolutionList resolutions = CameraSnap.GetResolutionList(this._Camera.Moniker);

            if (resolutions == null)return;

            int index_to_select = -1;

            for (int index = 0; index < resolutions.Count; index++)
            {
                comboBoxResolutionList.Items.Add(resolutions[index].ToString());

                if (resolutions[index].CompareTo(this._Camera.Resolution) == 0)
                {
                    index_to_select = index;
                }
            }

            // select current resolution
            if (index_to_select >= 0)
            {
                this.comboBoxResolutionList.SelectedIndex = index_to_select;
                this.comboBoxResolutionList.Enabled = true;
            }
        }



        private void FillCameraList()
        {
            comboBoxCameraList.Items.Clear();

            _CameraChoice.UpdateDeviceList();

            foreach (var camera_device in _CameraChoice.Devices)
            {
                comboBoxCameraList.Items.Add(camera_device.Name);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (this.pictureBox1.Image != null)
            {
                this.pictureBox1.Image.Dispose();
                this.pictureBox1.Image = null;
            }

            if (this._Camera == null) return;


            if (this._Camera.Ready())
            {
                Bitmap bitmap = null;
                try
                {
                    bitmap = this._Camera.SnapshotSourceImage();
                    this.pictureBox1.Image = bitmap;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, @"Error while getting a snapshot");
                    this.CloseCamera();
                }

                return;

            }
        }

    }
}
