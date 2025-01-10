using System;
using System.IO;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Microsoft.Win32;

using ScriptPortal.Vegas;  // "ScriptPortal.Vegas" for Magix Vegas Pro 14 or above, "Sony.Vegas" for Sony Vegas Pro 13 or below

namespace Test_Script
{
    public class TextMediaProperties
    {
        public Size MediaSize = new Size(1920, 1080);
        public double MediaMilliseconds = 5000;
        public string Text = "";
        public string FontName = "Arial";
        public double FontSize = 48;
        public bool FontItalic = false;
        public bool FontBold = false;
        public int TextAlign = 0;
        public Color TextColor = Color.White;
        public Color OutlineColor = Color.Black;
        public double OutlineWidth = 0;
        public double LocationX = 0.5;
        public double LocationY = 0.5;
        public double ScaleX = 1;
        public double ScaleY = 1;
        public int CenterType = 4;
        public double Tracking = 0;
        public double LineSpacing = 1;
        public bool ShadowEnable = false;
        public Color ShadowColor = Color.Black;
        public double ShadowX = 0.2;
        public double ShadowY = 0.2;
        public double ShadowBlur = 0.4;
    }

    public class TextMediaGenerator
    {
        public Vegas myVegas;
        
        public const string UID_PROTYPE_TITLER = "{53FC0B44-BD58-4716-A90F-3EB43168DE81}";
        public const string UID_TITLES_AND_TEXT = "{Svfx:com.vegascreativesoftware:titlesandtext}";
        public const string UID_TITLES_AND_TEXT_SONY = "{Svfx:com.sonycreativesoftware:titlesandtext}";

        public PlugInNode plugInProTypeTitler = null, plugInTitlesAndText = null;

        public static string RoamingPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

        public void Main(Vegas vegas)
        {
            myVegas = vegas;
            myVegas.ResumePlaybackOnScriptExit = true;

            plugInProTypeTitler = myVegas.Generators.FindChildByUniqueID(UID_PROTYPE_TITLER);
            plugInTitlesAndText = myVegas.Generators.FindChildByUniqueID(UID_TITLES_AND_TEXT)
                               ?? myVegas.Generators.FindChildByUniqueID(UID_TITLES_AND_TEXT_SONY);

            if (plugInProTypeTitler == null || plugInTitlesAndText == null)
            {
                MessageBox.Show("PlugIn Not Found!");
                return;
            }

            string baseXmlPath = Path.Combine(Path.GetDirectoryName(Script.File), "ProTypeTitlerBase.xml");

            List<VideoEvent> vEvents = new List<VideoEvent>();
            foreach (Track myTrack in myVegas.Project.Tracks)
            {
                if (myTrack.IsVideo())
                {
                    foreach (TrackEvent evnt in myTrack.Events)
                    {
                        if (evnt.Selected && evnt.ActiveTake != null && evnt.ActiveTake.Media != null && evnt.ActiveTake.Media.IsGenerated())
                        {
                            vEvents.Add((VideoEvent)evnt);
                        }
                    }
                }
            }

            Dictionary<Media, Media> mediaPairs = new Dictionary<Media, Media>();

            foreach (VideoEvent vEvent in vEvents)
            {
                Media newMedia = mediaPairs.ContainsKey(vEvent.ActiveTake.Media) ? mediaPairs[vEvent.ActiveTake.Media] : null;
                if (newMedia == null)
                {
                    Effect ef = vEvent.ActiveTake.Media.Generator;
                    if (ef.PlugIn == plugInTitlesAndText)
                    {
                        OFXEffect ofx = ef.OFXEffect;

                        RichTextBox rtb = new RichTextBox(){ Rtf = ((OFXStringParameter)ofx["Text"]).Value };
                        OFXDouble2D location = ((OFXDouble2DParameter)ofx["Location"]).Value;
                        double scale = ((OFXDoubleParameter)ofx["Scale"]).Value;
                        
                        TextMediaProperties properties = new TextMediaProperties()
                        {
                            MediaSize = vEvent.ActiveTake.Media.GetVideoStreamByIndex(0).Size,
                            MediaMilliseconds = vEvent.ActiveTake.Media.Length.ToMilliseconds(),
                            Text = rtb.Text.Replace("\r\n","&#xA;").Replace("\n","&#xA;"),
                            FontName = rtb.SelectionFont.Name,
                            FontSize = rtb.SelectionFont.Size,
                            FontItalic = rtb.SelectionFont.Italic,
                            FontBold = rtb.SelectionFont.Bold,
                            TextAlign = (int)rtb.SelectionAlignment,
                            TextColor = ConvertToColor(((OFXRGBAParameter)ofx["TextColor"]).Value),
                            OutlineColor = ConvertToColor(((OFXRGBAParameter)ofx["OutlineColor"]).Value),
                            OutlineWidth = ((OFXDoubleParameter)ofx["OutlineWidth"]).Value,
                            LocationX = location.X,
                            LocationY = location.Y,
                            ScaleX = scale,
                            ScaleY = scale,
                            CenterType = ((OFXChoiceParameter)ofx["Alignment"]).Value.Index,
                            Tracking = ((OFXDoubleParameter)ofx["Tracking"]).Value,
                            LineSpacing = ((OFXDoubleParameter)ofx["LineSpacing"]).Value,
                            ShadowEnable = ((OFXBooleanParameter)ofx["ShadowEnable"]).Value,
                            ShadowColor = ConvertToColor(((OFXRGBAParameter)ofx["ShadowColor"]).Value),
                            ShadowX = ((OFXDoubleParameter)ofx["ShadowOffsetX"]).Value,
                            ShadowY = ((OFXDoubleParameter)ofx["ShadowOffsetY"]).Value,
                            ShadowBlur = ((OFXDoubleParameter)ofx["ShadowBlur"]).Value
                        };

                        newMedia = GenerateProTypeTitlerMedia(myVegas.Project, properties, baseXmlPath);
                        mediaPairs.Add(vEvent.ActiveTake.Media, newMedia);
                    }
                }

                if (newMedia == null)
                {
                    continue;
                }

                vEvent.AddTake(newMedia.GetVideoStreamByIndex(0), true);
            }
        }

        public static Color ConvertToColor(OFXColor ofxColor)
        {
            return Color.FromArgb((int)(ofxColor.A * 255), (int)(ofxColor.R * 255), (int)(ofxColor.G * 255), (int)(ofxColor.B * 255));
        }

        public static string ConvertToString(Color color)
        {
            return string.Format("{0},{1},{2},{3}", color.A, color.R, color.G, color.B);
        }

        public static void SaveDxtEffectPreset(PlugInNode plugIn, string presetName, string xmlString)
        {
            if (plugIn == null || plugIn.IsOFX)
            {
                return;
            }

            RegistryKey myReg = Registry.CurrentUser.CreateSubKey(Path.Combine("Software", "DXTransform", "Presets", plugIn.UniqueID));
            string filePath = (string)myReg.GetValue(presetName) ?? Path.Combine(System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileMajorPart > 13 ? Path.Combine(RoamingPath, "VEGAS", "FX Presets") : Path.Combine(RoamingPath, "Sony", "VEGAS", "FX Presets"), plugIn.UniqueID, presetName + ".dxp");
            if (myReg.GetValue(presetName) == null)
            {
                myReg.SetValue(presetName, filePath);
            }
            byte[] data = Encoding.UTF8.GetBytes(xmlString);

            byte[] lengthBytes = BitConverter.GetBytes(data.Length);
            if (!BitConverter.IsLittleEndian)
            {
                Array.Reverse(lengthBytes);
            }

            data = InsertBytes(data, lengthBytes);

            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                fileStream.Write(data, 0, data.Length);
            }
        }

        public Media GenerateProTypeTitlerMedia(Project myProject, TextMediaProperties properties, string baseXmlPath)
        {
            string presetName = Path.GetFileNameWithoutExtension(baseXmlPath) + "Temp";
            string[] textAlignTypes = new string[] { "Near", "Far", "Center" };
            string[] centerTypes = new string[] { "NearBaseline", "CenterBaseline", "FarBaseline", "NearCenter", "Center", "FarCenter", "NearTop", "CenterTop", "FarTop", "CustomOffset" };
            string xmlString = string.Format
            (
                Encoding.UTF8.GetString(File.ReadAllBytes(baseXmlPath)),
                properties.Text,
                properties.FontName,
                properties.FontSize / 10,
                properties.FontItalic ? "Italic" : "Normal",
                properties.FontBold ? "Bold" : "Normal",
                textAlignTypes[properties.TextAlign],
                ConvertToString(properties.TextColor),
                ConvertToString(properties.OutlineColor),
                properties.OutlineWidth,
                string.Format
                (
                    "{0},{1},{2},{3},{4},{5}",
                    (properties.LocationX - 0.5) * 320 / 9,
                    (properties.LocationY - 0.5) * 20,
                    (properties.CenterType % 3) / 2.0,
                    (properties.CenterType / 3 % 3) / 2.0,
                    properties.ScaleX,
                    properties.ScaleY
                ),
                centerTypes[properties.CenterType > 5 ? 9 : (properties.CenterType + (properties.CenterType < 3 ? 6 : 0))],
                properties.Tracking,
                properties.LineSpacing,
                properties.ShadowEnable.ToString().ToLower(),
                ConvertToString(properties.ShadowColor),
                properties.ShadowX / 10,
                properties.ShadowY / 10,
                properties.ShadowBlur / 5
            );
            SaveDxtEffectPreset(plugInProTypeTitler, presetName, xmlString);
            Media myMedia = Media.CreateInstance(myProject, plugInProTypeTitler);
            myMedia.Generator.Preset = presetName;
            myMedia.GetVideoStreamByIndex(0).Size = properties.MediaSize;
            myMedia.Length = Timecode.FromMilliseconds(properties.MediaMilliseconds);
            return myMedia;
        }

        public static byte[] InsertBytes(byte[] data1, byte[] data2, int insertPos = 0, int extendLength = 0)
        {
            byte[] newData = new byte[data1.Length + data2.Length + extendLength];
            Array.Copy(data1, 0, newData, 0, insertPos);
            Array.Copy(data2, 0, newData, insertPos, data2.Length);
            Array.Copy(data1, insertPos, newData, data2.Length + insertPos, data1.Length - insertPos);
            return newData;
        }
    }
}

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    {
        Test_Script.TextMediaGenerator test = new Test_Script.TextMediaGenerator();
        test.Main(vegas);
    }
}