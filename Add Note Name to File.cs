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
                string newFilePath = ModifyFile(m.FilePath, noteName);
                if (!File.Exists(newFilePath) || Path.GetExtension(newFilePath).ToLower() == ".sfl")
                {
                    continue;
                }
                m.ReplaceWith(Media.CreateInstance(myVegas.Project, newFilePath));
            }
        }

        static string ModifyFile(string filePath, string pitch)
        {
            byte[] data = File.ReadAllBytes(filePath), sizeData = null;

            byte[] riffPrefix = { 0x52, 0x49, 0x46, 0x46 };
            bool isWav = true, isMp3 = false;
            for (int i = 0; i < riffPrefix.Length; i++)
            {
                if (data[i] != riffPrefix[i])
                {
                    isWav = false;
                }
            }

            string newPath = Path.Combine(Path.GetDirectoryName(filePath), string.Format("{0}_{1}{2}", Path.GetFileNameWithoutExtension(filePath), pitch, Path.GetExtension(filePath)));

            if (!isWav)
            {
                if (Path.GetExtension(filePath).ToLower() == ".mp3")
                {
                    isMp3 = true;
                    bool hasId3 = true;
                    byte[] id3Prefix = { 0x49, 0x44, 0x33 };
                    for (int i = 0; i < id3Prefix.Length; i++)
                    {
                        if (data[i] != id3Prefix[i])
                        {
                            hasId3 = false;
                            break;
                        }
                    }

                    if (!hasId3)
                    {
                        id3Prefix = new byte[] {
                            0x49, 0x44, 0x33, 0x03, 0x00, 0x00, // ID3v2.3
                            0x00, 0x00, 0x00, 0x00              // Size
                        };
                        data = InsertBytes(data, id3Prefix);
                    }

                    sizeData = ConvertIntToID3v2Size(ConvertID3v2Size(data) + 0x27);
                }

                else
                {
                    newPath = filePath + ".sfl";
                    bool hasRiff = false;
                    if (File.Exists(newPath))
                    {
                        data = File.ReadAllBytes(newPath);
                        hasRiff = true;
                        for (int i = 0; i < riffPrefix.Length; i++)
                        {
                            if (data[i] != riffPrefix[i])
                            {
                                hasRiff = false;
                                break;
                            }
                        }
                    }

                    byte[] nameBytes = System.Text.Encoding.Default.GetBytes(Path.GetFileName(filePath));

                    if (!hasRiff)
                    {
                        riffPrefix = new byte[] {
                            0x52, 0x49, 0x46, 0x46, 0x00, 0x00, 0x00, 0x00,
                            0x53, 0x46, 0x50, 0x4C, 0x53, 0x46, 0x50, 0x49,
                            0x00, 0x00, 0x00, 0x00
                        };
                        data = InsertBytes(nameBytes, riffPrefix, 0, 2 - (nameBytes.Length % 2));
                    }

                    byte[] nameLengthBytes = BitConverter.GetBytes((nameBytes.Length / 2 * 2) + 1);
                    Array.Copy(nameLengthBytes, 0, data, 16, nameLengthBytes.Length);
                }
            }

            byte[] acidPrefix = isMp3 ? new byte[]  {
                0x47, 0x45, 0x4F, 0x42, 0x00, 0x00, 0x00, 0x27,
                0x00, 0x00, 0x00, 0x00, 0x00, 0x53, 0x66, 0x41,
                0x63, 0x69, 0x64, 0x43, 0x68, 0x75, 0x6E, 0x6B,
                0x00
            } : new byte[] { 0x61, 0x63, 0x69, 0x64, 0x18, 0x00, 0x00, 0x00 };

            int position = FindBytes(data, acidPrefix);

            if (position != -1)
            {
                data[position + acidPrefix.Length + 4] = pitchValues[pitch];
                data[position + acidPrefix.Length] = 0x02; // Avoid problems with some FL Studio output files
            }
            else
            {
                byte[] acidData = {
                    0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x80,
                    0x00, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00,
                    0x04, 0x00, 0x04, 0x00, 0x00, 0x00, 0xF0, 0x42
                };
                acidData[4] = pitchValues[pitch];
                acidData = InsertBytes(acidData, acidPrefix);
                data = isMp3 ? InsertBytes(data, acidData, 10) : InsertBytes(acidData, data);
            }

            if (!isMp3)
            {
                sizeData = BitConverter.GetBytes(data.Length - 8);
            }
            Array.Copy(sizeData, 0, data, isMp3 ? 6 : 4, sizeData.Length);

            try
            {
                using (FileStream fileStream = new FileStream(newPath, FileMode.Create, FileAccess.Write))
                {
                    fileStream.Write(data, 0, data.Length);
                }
            }
            catch
            {
                
            }

            return newPath;
        }

        static int ConvertID3v2Size(byte[] bytes, int offset = 6)
        {
            int size = 0;
            size = (bytes[offset] & 0x7F) << 21;
            size |= (bytes[offset + 1] & 0x7F) << 14;
            size |= (bytes[offset + 2] & 0x7F) << 7;
            size |= (bytes[offset + 3] & 0x7F);
            return size;
        }

        static byte[] ConvertIntToID3v2Size(int size)
        {
            byte[] bytes = new byte[4];
            bytes[0] = (byte)((size >> 21) & 0x7F);
            bytes[1] = (byte)((size >> 14) & 0x7F);
            bytes[2] = (byte)((size >> 7) & 0x7F);
            bytes[3] = (byte)(size & 0x7F);
            return bytes;
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

        static byte[] InsertBytes(byte[] data1, byte[] data2, int insertPos = 0, int extendLength = 0)
        {
            byte[] newData = new byte[data1.Length + data2.Length + extendLength];
            Array.Copy(data1, 0, newData, 0, insertPos);
            Array.Copy(data2, 0, newData, insertPos, data2.Length);
            Array.Copy(data1, insertPos, newData, data2.Length + insertPos, data1.Length - insertPos);
            return newData;
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