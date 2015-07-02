/*  3dpBurner Image2Gcode. A Image to GCODE converter for GRBL based devices.
    This file is part of 3dpBurner Image2Gcode application.
   
    Copyright (C) 2015  Adrian V. J. (villamany) contact: villamany@gmail.com
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/
// Form 1 (Main form)

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using _3dpBurnerImage2Gcode.Properties;

namespace _3dpBurnerImage2Gcode
{
    public partial class Form1 : Form
    {
        private const string ver = "v0.1";
        private readonly char[] floatChars = {'.', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9'};
        private readonly char[] intChars = {'.', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9'};
        private Bitmap adjustedImage;
        private float coordX; // X
        private string coordXStr; // String formated X
        private float coordY; // Y
        private string coordYStr; // String formated Y
        private bool imageLoaded = false;
        private string imagePath;
        private decimal lastSz; // last 'S' value for compare
        private float lastValue; // Aux for apply processing to image only when a new value is detected
        private float lastX; // Last x/y  coords for compare
        private float lastY;
        private string line;
        private Bitmap originalImage;
        private int originalWidth;
        private string outputPath;

        private float ratio; // Used to lock the aspect ratio when the option is selected
        private StringCollection recentImages;
        private decimal sz; // S (or Z)
        private char szChar; // Use 'S' or 'Z' for test laser power
        private string szStr; // String formated S

        public Form1()
        {
            InitializeComponent();
        }

        // Save settings
        private void saveSettings()
        {
            try
            {
                string set;
                Settings.Default.autoZoom = autoZoomToolStripMenuItem.Checked;
                if (imperialinToolStripMenuItem.Checked) set = "imperial";
                else set = "metric";
                Settings.Default.units = set;
                Settings.Default.width = tbWidth.Text;
                Settings.Default.height = tbHeight.Text;
                Settings.Default.resolution = tbRes.Text;
                Settings.Default.minPower = tbLaserMin.Text;
                Settings.Default.maxPower = tbLaserMax.Text;
                Settings.Default.header = rtbPreGcode.Text;
                Settings.Default.footer = rtbPostGcode.Text;
                Settings.Default.feedrate = tbFeedRate.Text;
                if (rbUseZ.Checked) set = "Z";
                else set = "S";
                Settings.Default.powerCommand = set;
                Settings.Default.pattern = cbEngravingPattern.Text;
                Settings.Default.edgeLines = cbEdgeLines.Checked;
                Settings.Default.imagePath = imagePath;
                Settings.Default.outputPath = outputPath;
                Settings.Default.recentImages = recentImages;

                Settings.Default.Save();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error saving config: " + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Load settings
        private void loadSettings()
        {
            try
            {
                autoZoomToolStripMenuItem.Checked = Settings.Default.autoZoom;
                autoZoomToolStripMenuItem_Click(this, null);

                if (Settings.Default.units == "imperial")
                {
                    imperialinToolStripMenuItem.Checked = true;
                    imperialinToolStripMenuItem_Click(this, null);
                }
                else
                {
                    metricmmToolStripMenuItem.Checked = true;
                    metricmmToolStripMenuItem_Click(this, null);
                }
                tbWidth.Text = Settings.Default.width;
                tbHeight.Text = Settings.Default.height;
                tbRes.Text = Settings.Default.resolution;
                tbLaserMin.Text = Settings.Default.minPower;
                tbLaserMax.Text = Settings.Default.maxPower;
                rtbPreGcode.Text = Settings.Default.header;
                rtbPostGcode.Text = Settings.Default.footer;
                tbFeedRate.Text = Settings.Default.feedrate;
                if (Settings.Default.powerCommand == "Z")
                    rbUseZ.Checked = true;
                else rbUseS.Checked = true;
                cbEngravingPattern.Text = Settings.Default.pattern;
                cbEdgeLines.Checked = Settings.Default.edgeLines;
                imagePath = Settings.Default.imagePath;
                outputPath = Settings.Default.outputPath;
                recentImages = Settings.Default.recentImages;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error saving config: " + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // On form load
        private void Form1_Load(object sender, EventArgs e)
        {
            Text = "3dpBurner Image2Gcode " + ver;
            lblStatus.Text = "Done";
            loadSettings();

            autoZoomToolStripMenuItem_Click(this, null); // Set preview zoom mode

            refreshRecentImagesMenu();
        }

        // Interpolate a 8 bit grayscale value (0-255) between min,max
        // Changed to decimal for use with other controllers
        private decimal interpolate(decimal grayValue, decimal min, decimal max)
        {
            decimal dif = max - min;
            return (min + grayValue*dif)/255;
        }

        // Return true if char is a valid float digit, show eror message is not and return

        private bool checkDigitFloat(char ch)
        {
            if ((!floatChars.Contains(ch)) & (Convert.ToByte(ch) != 8) & (Convert.ToByte(ch) != 13))
                // is a 0-9 number or . decimal separator, backspace or enter
            {
                MessageBox.Show("Allowed chars are '0'-'9' and '.' as decimal separator", "Invalid value",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // Return true if char is a valid integer digit, show error message if not and return false

        private bool checkDigitInteger(char ch)
        {
            if ((!intChars.Contains(ch)) & (Convert.ToByte(ch) != 8) & (Convert.ToByte(ch) != 13))
                // is a 0-9 number or . decimal separator, backspace or enter
            {
                MessageBox.Show("Allowed chars are '0'-'9'", "Invalid value", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // Apply dithering to an image (Convert to 1 bit)
        private Bitmap imgDither(Bitmap input)
        {
            lblStatus.Text = "Dithering...";
            Refresh();
            var masks = new byte[] {0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01};
            var output = new Bitmap(input.Width, input.Height, PixelFormat.Format1bppIndexed);
            var data = new sbyte[input.Width, input.Height];
            BitmapData inputData = input.LockBits(new Rectangle(0, 0, input.Width, input.Height), ImageLockMode.ReadOnly,
                PixelFormat.Format24bppRgb);
            try
            {
                IntPtr scanLine = inputData.Scan0;
                var line = new byte[inputData.Stride];
                for (int y = 0; y < inputData.Height; y++, scanLine += inputData.Stride)
                {
                    Marshal.Copy(scanLine, line, 0, line.Length);
                    for (int x = 0; x < input.Width; x++)
                    {
                        data[x, y] = (sbyte) (64*(GetGreyLevel(line[x*3 + 2], line[x*3 + 1], line[x*3 + 0]) - 0.5));
                    }
                }
            }
            finally
            {
                input.UnlockBits(inputData);
            }
            BitmapData outputData = output.LockBits(new Rectangle(0, 0, output.Width, output.Height),
                ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed);
            try
            {
                IntPtr scanLine = outputData.Scan0;
                for (int y = 0; y < outputData.Height; y++, scanLine += outputData.Stride)
                {
                    var line = new byte[outputData.Stride];
                    for (int x = 0; x < input.Width; x++)
                    {
                        bool j = data[x, y] > 0;
                        if (j) line[x/8] |= masks[x%8];
                        var error = (sbyte) (data[x, y] - (j ? 32 : -32));
                        if (x < input.Width - 1) data[x + 1, y] += (sbyte) (7*error/16);
                        if (y < input.Height - 1)
                        {
                            if (x > 0) data[x - 1, y + 1] += (sbyte) (3*error/16);
                            data[x, y + 1] += (sbyte) (5*error/16);
                            if (x < input.Width - 1) data[x + 1, y + 1] += (sbyte) (1*error/16);
                        }
                    }
                    Marshal.Copy(line, 0, scanLine, outputData.Stride);
                }
            }
            finally
            {
                output.UnlockBits(outputData);
            }
            lblStatus.Text = "Done";
            Refresh();
            return (output);
        }

        private static double GetGreyLevel(byte r, byte g, byte b) // aux for dithering
        {
            return (r*0.299 + g*0.587 + b*0.114)/255;
        }

        // Adjust brightness contrast and gamma of an image      
        private Bitmap imgBalance(Bitmap img, int brigh, int cont, int gam)
        {
            lblStatus.Text = "Balancing...";
            Refresh();
            ImageAttributes imageAttributes;
            float brightness = (brigh/100.0f) + 1.0f;
            float contrast = (cont/100.0f) + 1.0f;
            float gamma = 1/(gam/100.0f);
            float adjustedBrightness = brightness - 1.0f;
            Bitmap output;
            // create matrix that will brighten and contrast the image
            float[][] ptsArray =
            {
                new[] {contrast, 0, 0, 0, 0}, // scale red
                new[] {0, contrast, 0, 0, 0}, // scale green
                new[] {0, 0, contrast, 0, 0}, // scale blue
                new[] {0, 0, 0, 1.0f, 0}, // don't scale alpha
                new[] {adjustedBrightness, adjustedBrightness, adjustedBrightness, 0, 1}
            };

            output = new Bitmap(img);
            imageAttributes = new ImageAttributes();
            imageAttributes.ClearColorMatrix();
            imageAttributes.SetColorMatrix(new ColorMatrix(ptsArray), ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
            imageAttributes.SetGamma(gamma, ColorAdjustType.Bitmap);
            Graphics g = Graphics.FromImage(output);
            g.DrawImage(output, new Rectangle(0, 0, output.Width, output.Height)
                , 0, 0, output.Width, output.Height,
                GraphicsUnit.Pixel, imageAttributes);
            lblStatus.Text = "Done";
            Refresh();
            return (output);
        }

        // Return a grayscale version of an image
        private Bitmap imgGrayscale(Bitmap original)
        {
            lblStatus.Text = "Grayscaling...";
            Refresh();
            var newBitmap = new Bitmap(original.Width, original.Height);
                // create a blank bitmap the same size as original
            Graphics g = Graphics.FromImage(newBitmap); // get a graphics object from the new image
            // create the grayscale ColorMatrix
            var colorMatrix = new ColorMatrix(
                new[]
                {
                    new[] {.299f, .299f, .299f, 0, 0},
                    new[] {.587f, .587f, .587f, 0, 0},
                    new[] {.114f, .114f, .114f, 0, 0},
                    new float[] {0, 0, 0, 1, 0},
                    new float[] {0, 0, 0, 0, 1}
                });
            var attributes = new ImageAttributes(); // create some image attributes
            attributes.SetColorMatrix(colorMatrix); // set the color matrix attribute

            // draw the original image on the new image using the grayscale color matrix
            g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height),
                0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);
            g.Dispose(); // dispose the Graphics object
            lblStatus.Text = "Done";
            Refresh();
            return (newBitmap);
        }

        // Return a inverted colors version of a image
        private Bitmap imgInvert(Bitmap original)
        {
            lblStatus.Text = "Inverting...";
            Refresh();
            var newBitmap = new Bitmap(original.Width, original.Height);
                // create a blank bitmap the same size as original
            Graphics g = Graphics.FromImage(newBitmap); // get a graphics object from the new image
            // create the grayscale ColorMatrix
            var colorMatrix = new ColorMatrix(
                new[]
                {
                    new float[] {-1, 0, 0, 0, 0},
                    new float[] {0, -1, 0, 0, 0},
                    new float[] {0, 0, -1, 0, 0},
                    new float[] {0, 0, 0, 1, 0},
                    new float[] {1, 1, 1, 0, 1}
                });
            var attributes = new ImageAttributes(); // create some image attributes
            attributes.SetColorMatrix(colorMatrix); // set the color matrix attribute

            // draw the original image on the new image using the grayscale color matrix
            g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height),
                0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);
            g.Dispose(); // dispose the Graphics object
            lblStatus.Text = "Done";
            Refresh();
            return (newBitmap);
        }

        // Resize image to desired width/height for gcode generation
        private Bitmap imgResize(Bitmap input, Int32 xSize, Int32 ySize)
        {
            // Resize
            Bitmap output;
            lblStatus.Text = "Resizing...";
            Refresh();
            output = new Bitmap(input, new Size(xSize, ySize));
            lblStatus.Text = "Done";
            Refresh();
            return (output);
        }

        // Invoked when the user input any value for image adjust
        private void userAdjust()
        {
            try
            {
                if (adjustedImage == null) return; // if no image, do nothing
                // Apply resize to original image
                Int32 xSize; // Total X pixels of resulting image for GCode generation
                Int32 ySize; // Total Y pixels of resulting image for GCode generation
                xSize =
                    Convert.ToInt32(float.Parse(tbWidth.Text, CultureInfo.InvariantCulture.NumberFormat)/
                                    float.Parse(tbRes.Text, CultureInfo.InvariantCulture.NumberFormat));
                ySize =
                    Convert.ToInt32(float.Parse(tbHeight.Text, CultureInfo.InvariantCulture.NumberFormat)/
                                    float.Parse(tbRes.Text, CultureInfo.InvariantCulture.NumberFormat));
                adjustedImage = imgResize(originalImage, xSize, ySize);
                // Apply balance to adjusted (resized) image
                adjustedImage = imgBalance(adjustedImage, tBarBrightness.Value, tBarContrast.Value, tBarGamma.Value);
                // Reset dithering to adjusted (resized and balanced) image
                cbDithering.Text = "GrayScale 8 bit";
                // Display image
                pictureBox1.Image = adjustedImage;
                // Set preview
                autoZoomToolStripMenuItem_Click(this, null);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error resizing/balancing image: " + e.Message, "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // Display a invalid values message
        private void invalidValue()
        {
            lblStatus.Text = "Invalid values! Check it.";
        }

        // Contrast adjusted by user
        private void tBarContrast_Scroll(object sender, EventArgs e)
        {
            lblContrast.Text = Convert.ToString(tBarContrast.Value);
            Refresh();
            userAdjust();
        }

        // Brightness adjusted by user
        private void tBarBrightness_Scroll(object sender, EventArgs e)
        {
            lblBrightness.Text = Convert.ToString(tBarBrightness.Value);
            Refresh();
            userAdjust();
        }

        // Gamma adjusted by user
        private void tBarGamma_Scroll(object sender, EventArgs e)
        {
            lblGamma.Text = Convert.ToString(tBarGamma.Value/100.0f);
            Refresh();
            userAdjust();
        }

        // Quick preview of the original image.
        private void btnCheckOrig_MouseDown(object sender, MouseEventArgs e)
        {
            if (adjustedImage == null) return; // if no image, do nothing
            if (!imageLoaded) return;
            lblStatus.Text = "Loading original image...";
            Refresh();
            pictureBox1.Image = originalImage;
            if (autoZoomToolStripMenuItem.Checked)
            {
                originalWidth = pictureBox1.Width;
                pictureBox1.Width = (originalWidth/2) - 5;
                pictureBox2.Width = (originalWidth/2) - 5;
                pictureBox2.Left = pictureBox1.Width + 10;
                pictureBox2.Image = adjustedImage;
            }
            lblStatus.Text = "Done";
        }

        // Reload the processed image after temporal preview of the original image
        private void btnCheckOrig_MouseUp(object sender, MouseEventArgs e)
        {
            if (adjustedImage == null) return; // if no image, do nothing
            if (!imageLoaded) return;
            pictureBox1.Image = adjustedImage;
            if (autoZoomToolStripMenuItem.Checked)
            {
                pictureBox1.Width = originalWidth;
                pictureBox2.Width = originalWidth;
                pictureBox2.Left = 0;
                pictureBox2.Image = null;
            }
        }

        // Check if a new image width has been confirmed by user, process it.
        private void widthChangedCheck()
        {
            try
            {
                if (adjustedImage == null) return; // if no image, do nothing
                float newValue = Convert.ToSingle(tbWidth.Text, CultureInfo.InvariantCulture.NumberFormat);
                    // Get the user input value           
                if (newValue == lastValue) return; // if not a new value do nothing     
                lastValue = Convert.ToSingle(tbWidth.Text, CultureInfo.InvariantCulture.NumberFormat);
                if (cbLockRatio.Checked)
                {
                    tbHeight.Text = Convert.ToString((newValue/ratio), CultureInfo.InvariantCulture.NumberFormat);
                }
                userAdjust();
            }
            catch
            {
                MessageBox.Show("Check width value.", "Invalid value", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Check if a new image height has been confirmed by user, process it.
        private void heightChangedCheck()
        {
            try
            {
                if (adjustedImage == null) return; // if no image, do nothing
                float newValue = Convert.ToSingle(tbHeight.Text, CultureInfo.InvariantCulture.NumberFormat);
                    // Get the user input value   
                if (newValue == lastValue) return; // if not is a new value do nothing
                lastValue = Convert.ToSingle(tbHeight.Text, CultureInfo.InvariantCulture.NumberFormat);
                if (cbLockRatio.Checked)
                {
                    tbWidth.Text = Convert.ToString((newValue*ratio), CultureInfo.InvariantCulture.NumberFormat);
                }
                userAdjust();
            }
            catch
            {
                MessageBox.Show("Check height value.", "Invalid value", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Check if a new image resolution has been confirmed by user, process it.
        private void resolutionChangedCheck()
        {
            try
            {
                if (adjustedImage == null) return; // if no image, do nothing
                float newValue = Convert.ToSingle(tbRes.Text, CultureInfo.InvariantCulture.NumberFormat);
                    // Get the user input value
                if (newValue == lastValue) return; // if not is a new value do nothing
                lastValue = Convert.ToSingle(tbRes.Text, CultureInfo.InvariantCulture.NumberFormat);
                userAdjust();
            }
            catch
            {
                MessageBox.Show("Check resolution value.", "Invalid value", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // CheckBox lockAspectRatio checked. Set as mandatory the user set width and calculate the height by using the original aspect ratio
        private void cbLockRatio_CheckedChanged(object sender, EventArgs e)
        {
            if (cbLockRatio.Checked)
            {
                tbHeight.Text =
                    Convert.ToString((Convert.ToSingle(tbWidth.Text, CultureInfo.InvariantCulture.NumberFormat)/ratio),
                        CultureInfo.InvariantCulture.NumberFormat); // Initialize y size
                if (adjustedImage == null) return; // if no image, do nothing
                userAdjust();
            }
        }

        // Width confirmed by user
        private void tbWidth_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Prevent any not allowed char
            if (!checkDigitFloat(e.KeyChar))
            {
                e.Handled = true; // Stop the character from being entered into the control since it is non-numerical.
                return;
            }

            if (e.KeyChar == Convert.ToChar(13))
            {
                widthChangedCheck();
            }
        }

        // Height confirmed by user by the enter key
        private void tbHeight_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Prevent any not allowed char
            if (!checkDigitFloat(e.KeyChar))
            {
                e.Handled = true; // Stop the character from being entered into the control since it is non-numerical.
                return;
            }
            if (e.KeyChar == Convert.ToChar(13))
            {
                heightChangedCheck();
            }
        }

        // Resolution confirmed by user by the enter key
        private void tbRes_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Prevent any not allowed char
            if (!checkDigitFloat(e.KeyChar))
            {
                e.Handled = true; // Stop the character from being entered into the control since it is non-numerical.
                return;
            }
            if (e.KeyChar == Convert.ToChar(13))
            {
                resolutionChangedCheck();
            }
        }

        // Width control leave focus. Check if new value
        private void tbWidth_Leave(object sender, EventArgs e)
        {
            widthChangedCheck();
        }

        // Height control leave focus. Check if new value
        private void tbHeight_Leave(object sender, EventArgs e)
        {
            heightChangedCheck();
        }

        // Resolution control leave focus. Check if new value
        private void tbRes_Leave(object sender, EventArgs e)
        {
            resolutionChangedCheck();
        }

        // Width control get focus
        private void tbWidth_Enter(object sender, EventArgs e)
        {
            try
            {
                lastValue = Convert.ToSingle(tbWidth.Text, CultureInfo.InvariantCulture.NumberFormat);
            }
            catch
            {
            }
        }

        // Height control get focus. Backup actual value for check changes.
        private void tbHeight_Enter(object sender, EventArgs e)
        {
            try
            {
                lastValue = Convert.ToSingle(tbHeight.Text, CultureInfo.InvariantCulture.NumberFormat);
            }
            catch
            {
            }
        }

        // Resolution control get focus. Backup actual value for check changes.
        private void tbRes_Enter(object sender, EventArgs e)
        {
            try
            {
                lastValue = Convert.ToSingle(tbRes.Text, CultureInfo.InvariantCulture.NumberFormat);
            }
            catch
            {
            }
        }

        // Generate a "minimalist" gcode line based on the actual and last coordinates and laser power
        // Changed sz and lastSz to decimal to work with Smoothieboard 0.0 to 1.0 power levels

        private void generateLine()
        {
            // Generate Gcode line
            // Changed to G1 formatted code to work with non-GRBL controllers
            line = "";
            if (coordX != lastX) // Add X coord to line if is different from previous             
            {
                coordXStr = string.Format(CultureInfo.InvariantCulture.NumberFormat, "{0:0.###}", coordX);
                line += "G1 X" + coordXStr;
            }
            if (coordY != lastY) // Add Y coord to line if is different from previous
            {
                coordYStr = string.Format(CultureInfo.InvariantCulture.NumberFormat, "{0:0.###}", coordY);
                line += "G1 Y" + coordYStr;
            }
            if (sz != lastSz) // Add power value to line if is diferent from previous
            {
                szStr = szChar + Convert.ToString(sz);
                line += szStr;
            }
        }

        // Generate button click
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (adjustedImage == null) return; // if no image, do nothing
            float resol = Convert.ToSingle(tbRes.Text, CultureInfo.InvariantCulture.NumberFormat);
                // Resolution (or laser spot size)
            float w = Convert.ToSingle(tbWidth.Text, CultureInfo.InvariantCulture.NumberFormat);
                // Get the user input value only for check for cancel if not valid         
            float h = Convert.ToSingle(tbHeight.Text, CultureInfo.InvariantCulture.NumberFormat);
                // Get the user input value only for check for cancel if not valid              

            if ((resol <= 0) | (adjustedImage.Width < 1) | (adjustedImage.Height < 1) | (w < 1) | (h < 1))
            {
                MessageBox.Show("Check width, height and resolution values.", "Invalid value", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }
            if (Convert.ToInt32(tbFeedRate.Text) < 1)
            {
                MessageBox.Show("Check feedrate value.", "Invalid value", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel) return;
            Int32 lin; // top/bottom pixel
            Int32 col; // Left/right pixel

            lblStatus.Text = "Generating file...";
            Refresh();

            List<string> fileLines;
            fileLines = new List<string>();
            // S or Z use as power command
            if (rbUseS.Checked) szChar = 'S';
            else szChar = 'Z';

            // first Gcode line
            line = "(Generated by 3dpBurner Image2Gcode " + ver + ")";
            fileLines.Add(line);
            line = "(@" + DateTime.Now.ToString("MMM/dd/yyy HH:mm:ss)");
            fileLines.Add(line);


            line = "M5\r"; // Make sure laser off
            fileLines.Add(line);

            // Add the pre-Gcode lines
            lastX = -1; // reset last positions
            lastY = -1;
            lastSz = -1;
            foreach (string s in rtbPreGcode.Lines)
            {
                fileLines.Add(s);
            }
            line = "G90\r"; // Absolute coordinates
            fileLines.Add(line);

            if (imperialinToolStripMenuItem.Checked) line = "G20\r"; // Imperial units
            else line = "G21\r"; // Metric units
            fileLines.Add(line);
            line = "F" + tbFeedRate.Text + "\r"; // Feedrate
            fileLines.Add(line);

            // Generate picture Gcode
            Int32 pixTot = adjustedImage.Width*adjustedImage.Height;
            Int32 pixBurned = 0;
            //////////////////////////////////////////////
            // Generate Gcode lines by Horizontal scanning
            //////////////////////////////////////////////
            if (cbEngravingPattern.Text == "Horizontal scanning")
            {
                // Goto rapid move to lef top corner
                line = "G0X0Y" +
                       string.Format(CultureInfo.InvariantCulture.NumberFormat, "{0:0.###}",
                           adjustedImage.Height*Convert.ToSingle(tbRes.Text, CultureInfo.InvariantCulture.NumberFormat));
                fileLines.Add(line);
                line = "G1\r"; // G1 mode
                fileLines.Add(line);
                line = "M3\r"; // Switch laser on
                fileLines.Add(line);

                // Start image
                lin = adjustedImage.Height - 1; // top tile
                col = 0; // Left pixel
                while (lin >= 0)
                {
                    // Y coordinate
                    coordY = resol*lin;
                    while (col < adjustedImage.Width) // From left to right
                    {
                        // X coordinate
                        coordX = resol*col;
                        // Power value
                        Color cl = adjustedImage.GetPixel(col, (adjustedImage.Height - 1) - lin); // Get pixel color
                        sz = 255 - cl.R;
                        sz = interpolate(sz, Convert.ToDecimal(tbLaserMin.Text), Convert.ToDecimal(tbLaserMax.Text));
                        generateLine();
                        pixBurned++;
                        //adjustedImage.SetPixel(col, (adjustedImage.Height-1)-lin, Color.Red);
                        //pictureBox1.Image = adjustedImage;
                        //Refresh();

                        if (!string.IsNullOrEmpty(line)) fileLines.Add(line);
                        lastX = coordX;
                        lastY = coordY;
                        lastSz = sz;
                        col++;
                    }
                    col--;
                    lin--;
                    coordY = resol*lin;
                    while ((col >= 0) & (lin >= 0)) // From right to left
                    {
                        // X coordinate
                        coordX = resol*col;
                        // Power value
                        Color cl = adjustedImage.GetPixel(col, (adjustedImage.Height - 1) - lin); // Get pixel color
                        sz = 255 - cl.R;
                        sz = interpolate(sz, Convert.ToDecimal(tbLaserMin.Text), Convert.ToDecimal(tbLaserMax.Text));
                        generateLine();
                        pixBurned++;
                        //adjustedImage.SetPixel(col, (adjustedImage.Height-1)-lin, Color.Red);
                        //pictureBox1.Image = adjustedImage;
                        //Refresh();

                        if (!string.IsNullOrEmpty(line)) fileLines.Add(line);
                        lastX = coordX;
                        lastY = coordY;
                        lastSz = sz;
                        col--;
                    }
                    col++;
                    lin--;
                    lblStatus.Text = "Generating file... " + Convert.ToString((pixBurned*100)/pixTot) + "%";
                    Refresh();
                }
            }
                //////////////////////////////////////////////
                // Generate Gcode lines by Diagonal scanning
                //////////////////////////////////////////////
            else
            {
                // Goto rapid move to lef top corner
                line = "G0X0Y0";
                fileLines.Add(line);
                line = "G1\r"; // G1 mode
                fileLines.Add(line);
                line = "M3\r"; // Switch laser on
                fileLines.Add(line);

                // Start image
                col = 0;
                lin = 0;
                while ((col < adjustedImage.Width) | (lin < adjustedImage.Height))
                {
                    while ((col < adjustedImage.Width) & (lin >= 0))
                    {
                        // Y coordinate
                        coordY = resol*lin;
                        // X coordinate
                        coordX = resol*col;

                        // Power value
                        Color cl = adjustedImage.GetPixel(col, (adjustedImage.Height - 1) - lin); // Get pixel color
                        sz = 255 - cl.R;
                        sz = interpolate(sz, Convert.ToDecimal(tbLaserMin.Text), Convert.ToDecimal(tbLaserMax.Text));

                        generateLine();
                        pixBurned++;

                        //adjustedImage.SetPixel(col, (adjustedImage.Height-1)-lin, Color.Red);
                        //pictureBox1.Image = adjustedImage;
                        //Refresh();

                        if (!string.IsNullOrEmpty(line)) fileLines.Add(line);
                        lastX = coordX;
                        lastY = coordY;
                        lastSz = sz;

                        col++;
                        lin--;
                    }
                    col--;
                    lin++;

                    if (col >= adjustedImage.Width - 1) lin++;
                    else col++;
                    while ((col >= 0) & (lin < adjustedImage.Height))
                    {
                        // Y coordinate
                        coordY = resol*lin;
                        // X coordinate
                        coordX = resol*col;

                        // Power value
                        Color cl = adjustedImage.GetPixel(col, (adjustedImage.Height - 1) - lin); // Get pixel color
                        sz = 255 - cl.R;
                        sz = interpolate(sz, Convert.ToDecimal(tbLaserMin.Text), Convert.ToDecimal(tbLaserMax.Text));

                        generateLine();
                        pixBurned++;

                        //adjustedImage.SetPixel(col, (adjustedImage.Height-1)-lin, Color.Red);
                        //pictureBox1.Image = adjustedImage;
                        // Refresh();

                        if (!string.IsNullOrEmpty(line)) fileLines.Add(line);
                        lastX = coordX;
                        lastY = coordY;
                        lastSz = sz;

                        col--;
                        lin++;
                    }
                    col++;
                    lin--;
                    if (lin >= adjustedImage.Height - 1) col++;
                    else lin++;
                    lblStatus.Text = "Generating file... " + Convert.ToString((pixBurned*100)/pixTot) + "%";
                    Refresh();
                }
            }
            // Edge lines
            if (cbEdgeLines.Checked)
            {
                line = "M5\r";
                fileLines.Add(line);
                line = "G0X0Y0\r";
                fileLines.Add(line);
                line = "M3S" + tbLaserMax.Text + "\r";
                fileLines.Add(line);
                line = "G1X0Y" +
                       string.Format(CultureInfo.InvariantCulture.NumberFormat, "{0:0.###}",
                           (adjustedImage.Height - 1)*resol) + "\r";
                fileLines.Add(line);
                line = "G1X" +
                       string.Format(CultureInfo.InvariantCulture.NumberFormat, "{0:0.###}",
                           (adjustedImage.Width - 1)*resol) + "Y" +
                       string.Format(CultureInfo.InvariantCulture.NumberFormat, "{0:0.###}",
                           (adjustedImage.Height - 1)*resol) + "\r";
                fileLines.Add(line);
                line = "G1X" +
                       string.Format(CultureInfo.InvariantCulture.NumberFormat, "{0:0.###}",
                           (adjustedImage.Width - 1)*resol) + "Y0\r";
                fileLines.Add(line);
                line = "G1X0Y0\r";
                fileLines.Add(line);
            }
            // Switch laser off
            line = "M5\r"; //G1 mode
            fileLines.Add(line);

            // Add the post-Gcode 
            foreach (string s in rtbPostGcode.Lines)
            {
                fileLines.Add(s);
            }
            lblStatus.Text = "Saving file...";
            Refresh();
            // Save file
            File.WriteAllLines(saveFileDialog1.FileName, fileLines);
            int outputPathIndex = saveFileDialog1.FileName.LastIndexOf(@"\");
            outputPath = saveFileDialog1.FileName.Substring(0, outputPathIndex);
            saveSettings();
            saveFileDialog1.InitialDirectory = outputPath;
            lblStatus.Text = "Done (" + Convert.ToString(pixBurned) + "/" + Convert.ToString(pixTot) + ")";
        }

        // Horizontal mirroing
        private void btnHorizMirror_Click(object sender, EventArgs e)
        {
            if (adjustedImage == null) return; // if no image, do nothing
            lblStatus.Text = "Mirroring...";
            Refresh();
            adjustedImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
            originalImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
            pictureBox1.Image = adjustedImage;
            lblStatus.Text = "Done";
        }

        // Vertical mirroing
        private void btnVertMirror_Click(object sender, EventArgs e)
        {
            if (adjustedImage == null) return; // if no image, do nothing
            lblStatus.Text = "Mirroring...";
            Refresh();
            adjustedImage.RotateFlip(RotateFlipType.RotateNoneFlipY);
            originalImage.RotateFlip(RotateFlipType.RotateNoneFlipY);
            pictureBox1.Image = adjustedImage;
            lblStatus.Text = "Done";
        }

        // Rotate right
        private void btnRotateRight_Click(object sender, EventArgs e)
        {
            if (adjustedImage == null) return; // if no image, do nothing
            lblStatus.Text = "Rotating...";
            Refresh();
            adjustedImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
            originalImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
            ratio = 1/ratio;
            string s = tbHeight.Text;
            tbHeight.Text = tbWidth.Text;
            tbWidth.Text = s;
            pictureBox1.Image = adjustedImage;
            autoZoomToolStripMenuItem_Click(this, null);
            lblStatus.Text = "Done";
        }

        // Rotate left
        private void btnRotateLeft_Click(object sender, EventArgs e)
        {
            if (adjustedImage == null) return; // if no image, do nothing
            lblStatus.Text = "Rotating...";
            Refresh();
            adjustedImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
            originalImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
            ratio = 1/ratio;
            string s = tbHeight.Text;
            tbHeight.Text = tbWidth.Text;
            tbWidth.Text = s;
            pictureBox1.Image = adjustedImage;
            autoZoomToolStripMenuItem_Click(this, null);
            lblStatus.Text = "Done";
        }

        // Invert image color
        private void btnInvert_Click(object sender, EventArgs e)
        {
            if (adjustedImage == null) return; // if no image, do nothing
            adjustedImage = imgInvert(adjustedImage);
            originalImage = imgInvert(originalImage);
            pictureBox1.Image = adjustedImage;
        }

        private void cbDithering_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (adjustedImage == null) return; // if no image, do nothing
            if (cbDithering.Text == "Dithering FS 1 bit")
            {
                lblStatus.Text = "Dithering...";
                adjustedImage = imgDither(adjustedImage);
                pictureBox1.Image = adjustedImage;
                lblStatus.Text = "Done";
            }
            else
                userAdjust();
        }

        // Feedrate Text changed
        private void tbFeedRate_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Prevent any not allowed char
            if (!checkDigitInteger(e.KeyChar))
            {
                e.Handled = true; // Stop the character from being entered into the control since it is non-numerical.
            }
        }

        // Metric units selected
        private void metricmmToolStripMenuItem_Click(object sender, EventArgs e)
        {
            imperialinToolStripMenuItem.Checked = false;
            gbDimensions.Text = "Output (mm)";
            lblFeedRateUnits.Text = "mm/min";
        }

        // Imperial unitsSelected
        private void imperialinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            metricmmToolStripMenuItem.Checked = false;
            gbDimensions.Text = "Output (in)";
            lblFeedRateUnits.Text = "in/min";
        }

        // About dialog
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frmAbout = new Form2();
            frmAbout.ShowDialog();
        }

        // Preview AutoZoom
        private void autoZoomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (autoZoomToolStripMenuItem.Checked)
            {
                pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                pictureBox1.Width = panel1.Width;
                pictureBox1.Height = panel1.Height;
                pictureBox1.Top = 0;
                pictureBox1.Left = 0;
                pictureBox2.SizeMode = PictureBoxSizeMode.Zoom;
                pictureBox2.Width = panel1.Width;
                pictureBox2.Height = panel1.Height;
                pictureBox2.Top = 0;
                pictureBox2.Left = 0;
            }
            else
            {
                pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
                if (pictureBox1.Width > panel1.Width) pictureBox1.Left = 0;
                else pictureBox1.Left = (panel1.Width/2) - (pictureBox1.Width/2);
                if (pictureBox1.Height > panel1.Height) pictureBox1.Top = 0;
                else pictureBox1.Top = (panel1.Height/2) - (pictureBox1.Height/2);
            }
        }

        // Laser Min keyPress
        private void tbLaserMin_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Prevent any not allowed char
            if (!checkDigitFloat(e.KeyChar))
            {
                e.Handled = true; // Stop the character from being entered into the control since it is non-numerical.
            }
        }

        // Laser Max keyPress
        private void tbLaserMax_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Prevent any not allowed char
            if (!checkDigitFloat(e.KeyChar))
            {
                e.Handled = true; // Stop the character from being entered into the control since it is non-numerical.
            }
        }

        // OpenFile, save picture grayscaled to originalImage and save the original aspect ratio to ratio
        // Save directory path for opening images
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel) return; // if no image, do nothing
            if (!File.Exists(openFileDialog1.FileName))
            {
                MessageBox.Show("File does not exist");
                return;
            }
            openFileAndConvert(openFileDialog1.FileName);
        }

        private void openFileAndConvert(string filePath)
        {
            try
            {
                imageLoaded = true;
                lblStatus.Text = "Opening file...";
                int pathEndIndex = filePath.LastIndexOf(@"\");
                int extensionIndex = filePath.LastIndexOf(@".");
                int fileNameLength = extensionIndex - pathEndIndex - 1;
                imagePath = filePath.Substring(0, pathEndIndex);
                saveSettings();
                openFileDialog1.InitialDirectory = imagePath;
                saveFileDialog1.FileName = filePath.Substring(pathEndIndex + 1, fileNameLength);
                addToRecentImages(filePath);
                Refresh();
                tBarBrightness.Value = 0;
                tBarContrast.Value = 0;
                tBarGamma.Value = 100;
                lblBrightness.Text = Convert.ToString(tBarBrightness.Value);
                lblContrast.Text = Convert.ToString(tBarContrast.Value);
                lblGamma.Text = Convert.ToString(tBarGamma.Value/100.0f);
                originalImage = new Bitmap(Image.FromFile(filePath));
                originalImage = imgGrayscale(originalImage);
                adjustedImage = new Bitmap(originalImage);
                ratio = (originalImage.Width + 0.0f)/originalImage.Height; // Save ratio for future use if needled
                if (cbLockRatio.Checked)
                    tbHeight.Text = Convert.ToString((Convert.ToSingle(tbWidth.Text)/ratio),
                        CultureInfo.InvariantCulture.NumberFormat); // Initialize y size
                userAdjust();
                lblStatus.Text = "Done";
            }
            catch (Exception err)
            {
                MessageBox.Show("Error opening file: " + err.Message, "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void addToRecentImages(string filePath)
        {
            if (recentImages.Contains(filePath)) return;

            if (recentImages.Count < 20)
            {
                recentImages.Insert(0, filePath);
            }
            else
            {
                recentImages.RemoveAt(19);
                recentImages.Insert(0, filePath);
            }
            saveSettings();
            refreshRecentImagesMenu();
        }

        private void refreshRecentImagesMenu()
        {
            recentImagesToolStripMenuItem.DropDownItems.Clear();
            ToolStripItem item;
            foreach (string path in recentImages)
            {
                item = recentImagesToolStripMenuItem.DropDownItems.Add(path);
                item.Click += OnRecentFileClick;
            }
        }

        private void OnRecentFileClick(object sender, EventArgs e)
        {
            openFileAndConvert(sender.ToString());
        }

        // Exit Menu
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        // On form closing
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            saveSettings();
        }

        private void resetButton_Click(object sender, EventArgs e)
        {
            tBarBrightness.Value = 0;
            lblBrightness.Text = Convert.ToString(tBarBrightness.Value);
            tBarContrast.Value = 0;
            lblContrast.Text = Convert.ToString(tBarContrast.Value);
            tBarGamma.Value = 100;
            lblGamma.Text = Convert.ToString(tBarGamma.Value/100.0f);
            Refresh();
            userAdjust();
        }
    }
}