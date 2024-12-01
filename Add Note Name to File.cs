using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using ScriptPortal.Vegas;  // "ScriptPortal.Vegas" for Magix Vegas Pro 14 or above, "Sony.Vegas" for Sony Vegas Pro 13 or below

namespace Test_Script
{
    public class Class
    {
        public Vegas myVegas;

        static Dictionary<string, byte> pitchValues = new Dictionary<string, byte>
        {
            { "C", 0x30 }, { "C#", 0x31 }, { "Db", 0x31 }, { "D", 0x32 }, { "D#", 0x33 }, { "Eb", 0x33 },
            { "E", 0x34 }, { "F", 0x35 }, { "F#", 0x36 }, { "Gb", 0x36 }, { "G", 0x37 }, { "G#", 0x38 },
            { "Ab", 0x38 }, { "A", 0x39 }, { "A#", 0x3A }, { "Bb", 0x3A }, { "B", 0x3B }
        };

        public void Main(Vegas vegas)
        {
            myVegas = vegas;
            myVegas.ResumePlaybackOnScriptExit = true;
            List<Media> mediaList = new List<Media>();
            foreach (Track myTrack in myVegas.Project.Tracks)
            {
                if (myTrack.IsAudio())
                {
                    foreach (TrackEvent evnt in myTrack.Events)
                    {
                        if (evnt.Selected && evnt.ActiveTake != null && evnt.ActiveTake.Media != null && !mediaList.Contains(evnt.ActiveTake.Media))
                        {
                            mediaList.Add(evnt.ActiveTake.Media);
                        }
                    }
                }
            }

            if (mediaList.Count == 0)
            {
                return;
            }

            string noteName = GetBaseNoteName();
            if (string.IsNullOrEmpty(noteName) || !pitchValues.ContainsKey(noteName))
            {
                return;
            }

            foreach (Media m in mediaList)
            {
                string newFilePath = ModifyWavFile(m.FilePath, noteName);
                if (!File.Exists(newFilePath))
                {
                    continue;
                }
                m.ReplaceWith(Media.CreateInstance(myVegas.Project, newFilePath));
            }
            myVegas.Project.MediaPool.OpenAllMedia();
        }

        static string ModifyWavFile(string filePath, string pitch)
        {
            byte[] data = File.ReadAllBytes(filePath);

            byte[] riffPrefix = { 0x52, 0x49, 0x46, 0x46 };
            bool isWav = true;
            for (int i = 0; i < riffPrefix.Length; i++)
            {
                if (data[i] != riffPrefix[i])
                {
                    isWav = false;
                }
            }

            string sflPath = filePath + ".sfl";
            if (!isWav)
            {
                bool hasRiff = false;
                if (File.Exists(sflPath))
                {
                    data = File.ReadAllBytes(sflPath);
                    hasRiff = true;
                    for (int i = 0; i < riffPrefix.Length; i++)
                    {
                        if (data[i] != riffPrefix[i])
                        {
                            hasRiff = false;
                        }
                    }
                }

                byte[] nameBytes = System.Text.Encoding.Default.GetBytes(Path.GetFileName(filePath));

                if (!hasRiff)
                {
                    riffPrefix = new byte[] { 0x52, 0x49, 0x46, 0x46, 0x00, 0x00, 0x00, 0x00,
                                              0x53, 0x46, 0x50, 0x4C, 0x53, 0x46, 0x50, 0x49,
                                              0x00, 0x00, 0x00, 0x00 };
                    data = new byte[riffPrefix.Length + (nameBytes.Length / 2 + 1) * 2];
                    Array.Copy(riffPrefix, 0, data, 0, riffPrefix.Length);
                    Array.Copy(nameBytes, 0, data, riffPrefix.Length, nameBytes.Length);
                }

                byte[] nameLengthBytes = BitConverter.GetBytes((nameBytes.Length / 2 * 2) + 1);
                Array.Copy(nameLengthBytes, 0, data, 16, nameLengthBytes.Length);
            }

            byte[] targetData = { 0x61, 0x63, 0x69, 0x64, 0x18, 0x00, 0x00, 0x00 };
            int position = FindBytes(data, targetData);

            if (position != -1)
            {
                data[position + 12] = pitchValues[pitch];
                data[position + 8] = 0x02;
            }
            else
            {
                byte[] additionalData = {
                    0x61, 0x63, 0x69, 0x64, 0x18, 0x00, 0x00, 0x00,
                    0x02, 0x00, 0x00, 0x00, 0x30, 0x00, 0x00, 0x80,
                    0x00, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00,
                    0x04, 0x00, 0x04, 0x00, 0x00, 0x00, 0xF0, 0x42
                };
                additionalData[12] = pitchValues[pitch];
                byte[] newData = new byte[data.Length + additionalData.Length];
                Array.Copy(data, 0, newData, 0, data.Length);
                Array.Copy(additionalData, 0, newData, data.Length, additionalData.Length);
                data = newData;
            }

            byte[] riffData = BitConverter.GetBytes(data.Length - 8);
            Array.Copy(riffData, 0, data, 4, riffData.Length);

            string newPath = isWav ? Path.Combine(Path.GetDirectoryName(filePath), string.Format("{0}_{1}{2}", Path.GetFileNameWithoutExtension(filePath), pitch, Path.GetExtension(filePath))) : filePath;

            try
            {
                using (FileStream fileStream = new FileStream(isWav ? newPath : sflPath, FileMode.Create, FileAccess.Write))
                {
                    fileStream.Write(data, 0, data.Length);
                }
            }
            catch
            {
                
            }

            return newPath;
        }

        static int FindBytes(byte[] array, byte[] pattern)
        {
            int patternLength = pattern.Length;
            int totalLength = array.Length - patternLength + 1;
            for (int i = 0; i < totalLength; i++)
            {
                bool found = true;
                for (int j = 0; j < patternLength; j++)
                {
                    if (array[i + j] != pattern[j])
                    {
                        found = false;
                        break;
                    }
                }
                if (found)
                {
                    return i;
                }
            }
            return -1;
        }

        static string GetBaseNoteName()
        {
            Form form = new Form
            {
                ShowInTaskbar = false,
                AutoSize = true,
                BackColor = Color.FromArgb(45, 45, 45),
                ForeColor = Color.FromArgb(200, 200, 200),
                Font = new Font("Arial", 9),
                Text = "Note",
                FormBorderStyle = FormBorderStyle.FixedToolWindow,
                StartPosition = FormStartPosition.CenterScreen,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            Panel p = new Panel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            form.Controls.Add(p);

            TableLayoutPanel l = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                GrowStyle = TableLayoutPanelGrowStyle.AddRows,
                ColumnCount = 1
            };
            p.Controls.Add(l);

            ComboBox cb = new ComboBox()
            {
                DataSource = new string[]{ "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B" },
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(6, 12, 6, 9),
                Dock = DockStyle.Fill
            };
            l.Controls.Add(cb);

            Button ok = new Button
            {
                Font = new Font("Arial", 8),
                Text = "OK",
                Dock = DockStyle.Fill,
                Margin = new Padding(24, 4, 24, 6),
                DialogResult = DialogResult.OK
            };
            l.Controls.Add(ok);
            form.AcceptButton = ok;

            string value = null;
            if (form.ShowDialog() == DialogResult.OK)
            {
                value = cb.SelectedItem.ToString();
            }
            return value;
        }
    }
}

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    {
        Test_Script.Class test = new Test_Script.Class();
        test.Main(vegas);
    }
}