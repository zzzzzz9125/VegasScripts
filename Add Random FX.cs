using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using ScriptPortal.Vegas;  // "ScriptPortal.Vegas" for Magix Vegas Pro 14 or above, "Sony.Vegas" for Sony Vegas Pro 13 or below


namespace AddRandomFX
{
    public class AddRandomFXClass
    {
        // Here're the default settings that you can modify before running the script.
        static readonly bool DO_SHOW_WINDOW = true; // true: Show settings window before running the script; false: Use the default settings directly.

        static int FxCount = 1; // FX count for each target object.

        static int objectChoice = 0; // 0: Events, 1: Tracks, 2: Both

        // Regular expression. Ignore Case. Use ";" to separate. For exsample: "Color Grading;Stabiliz[e(ation)]".
        // If you need to check the Video FX list, please refer to: https://zzzzzz9125.github.io/VegTips/videofxlist
        static string blackWhiteListString_Name = "";

        // Regular expression. Ignore Case. Use ";" to separate. For exsample: "colorgrading;stabiliz[e(ation)]".
        // If you need to check the Video FX list, please refer to: https://zzzzzz9125.github.io/VegTips/videofxlist
        static string blackWhiteListString_UID = "colorgrading;stabiliz[e(ation)]";

        static bool blackWhiteList_Name = true; // true: Black List, false: White List.

        static bool blackWhiteList_UID = true; // true: Black List, false: White List.

        static bool randomPresets = true; // true: Random Presets, false: No Preset.


        public Vegas myVegas;
        static readonly Random rng = new Random();

        public void Main(Vegas vegas)
        {
            myVegas = vegas;
            myVegas.ResumePlaybackOnScriptExit = true;

            List<PlugInNode> l = new List<PlugInNode>();
            GetPluginNode(l, myVegas.VideoFX);

            if (DO_SHOW_WINDOW && !ShowWindow())
            {
                return;
            }

            List<Effects> effectsList = new List<Effects>();

            foreach (Track myTrack in myVegas.Project.Tracks)
            {
                if (myTrack.IsVideo())
                {
                    if (myTrack.Selected && (objectChoice == 1 || objectChoice == 2))
                    {
                        effectsList.Add(myTrack.Effects);
                    }
                    if (objectChoice == 0 || objectChoice == 2)
                    {
                        foreach (TrackEvent ev in myTrack.Events)
                        {
                            if (ev.Selected)
                            {
                                effectsList.Add((ev as VideoEvent).Effects);
                            }
                        }
                    }
                }
            }

            if (effectsList.Count == 0)
            {
                return;
            }

            List<PlugInNode> filtered_NameOnly = new List<PlugInNode>(), filtered = new List<PlugInNode>();
            

            if (blackWhiteList_Name && string.IsNullOrEmpty(blackWhiteListString_Name))
            {
                filtered_NameOnly.AddRange(l);
            }
            else
            {
                foreach (string str in blackWhiteListString_Name.Split(';'))
                {
                    if (string.IsNullOrEmpty(str))
                    {
                        continue;
                    }
                    filtered_NameOnly.AddRange(l.FindAll(node => Regex.IsMatch(node.Name, str, RegexOptions.IgnoreCase) != blackWhiteList_Name));
                }
            }

            if (blackWhiteList_UID && string.IsNullOrEmpty(blackWhiteListString_UID))
            {
                filtered.AddRange(filtered_NameOnly);
            }
            else
            {
                foreach (string str in blackWhiteListString_UID.Split(';'))
                {
                    if (string.IsNullOrEmpty(str))
                    {
                        continue;
                    }
                    filtered.AddRange(filtered_NameOnly.FindAll(node => Regex.IsMatch(node.UniqueID, str, RegexOptions.IgnoreCase) != blackWhiteList_UID));
                }
            }

            if (filtered.Count == 0)
            {
                MessageBox.Show(myVegas.MainWindow, "No FX match the filter criteria.", "Add Random FX", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (Effects effects in effectsList)
            {
                int effectsCount = effects.Count;
                do
                {
                    try
                    {
                        PlugInNode node = GetRandomElement(filtered);

                        Effect ef = new Effect(node);

                        effects.Add(ef);

                        if (randomPresets)
                        {
                            if (ef.Presets.Count > 0)
                            {
                                ef.Preset = GetRandomElement(ef.Presets).Name;
                            }
                        }
                    }
                    catch { }
                } while (effects.Count - effectsCount < FxCount);
            }
        }

        public void GetPluginNode(List<PlugInNode> list, PlugInNode node)
        {
            foreach(PlugInNode p in node)
            {
                if (p.IsContainer)
                {
                    GetPluginNode(list, p);
                }
                else
                {
                    bool repeated = false;
                    foreach (PlugInNode tmp in list)
                    {
                        if (tmp.UniqueID == p.UniqueID)
                        {
                            repeated = true;
                        }
                    }
                    if (!repeated)
                    {
                        list.Add(p);
                    }
                }
            }
        }

        public T GetRandomElement<T>(IList<T> list)
        {
            if (list == null || list.Count == 0)
            {
                return default(T);
            }

            int index = rng.Next(0, list.Count);
            return list[index];
        }

        public bool ShowWindow()
        {
            Form form = new Form
            {
                ShowInTaskbar = false,
                AutoSize = true,
                BackColor = Color.FromArgb(45, 45, 45),
                ForeColor = Color.FromArgb(200, 200, 200),
                Font = new Font("Arial", 9),
                Text = "Add Random FX",
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
                ColumnCount = 2
            };
            p.Controls.Add(l);

            ToolTip tt = new ToolTip();

            Label label = new Label
            {
                Margin = new Padding(12, 9, 6, 6),
                Text = "FX Count",
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = true
            };
            l.Controls.Add(label);

            TextBox fxCountTextBox = new TextBox
            {
                AutoSize = true,
                Margin = new Padding(6, 6, 6, 6),
                Text = FxCount.ToString(),
                Dock = DockStyle.Fill
            };
            l.Controls.Add(fxCountTextBox);

            tt.SetToolTip(fxCountTextBox, "FX count for each target object.");

            fxCountTextBox.MouseWheel += delegate (object o, MouseEventArgs e)
            {
                TextBox tb = o as TextBox;
                int tmp;
                if (int.TryParse(tb.Text, out tmp))
                {
                    tmp += e.Delta > 0 ? 1 : -1;
                    if (tmp > 0)
                    {
                        tb.Text = tmp.ToString();
                    }
                }
            };

            label = new Label
            {
                Margin = new Padding(12, 9, 6, 6),
                Text = "To Selected",
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = true
            };
            l.Controls.Add(label);

            ComboBox cb = new ComboBox()
            {
                DataSource = new string[] { "Events", "Tracks", "Both" },
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(6, 6, 6, 6),
                Dock = DockStyle.Fill
            };
            l.Controls.Add(cb);

            tt.SetToolTip(cb, "Select the target object type.");

            form.Load += delegate (object sender, EventArgs e)
            {
                cb.SelectedIndex = objectChoice;
            };

            Button blackWhiteListButton_Name = new Button
            {
                Text = blackWhiteList_Name ? "Name Black List" : "Name White List",
                Margin = new Padding(0, 3, 0, 3),
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                Anchor = AnchorStyles.None
            };
            l.Controls.Add(blackWhiteListButton_Name);
            blackWhiteListButton_Name.FlatAppearance.BorderSize = 0;

            tt.SetToolTip(blackWhiteListButton_Name, "Click to toggle between Name Black List and Name White List.");

            blackWhiteListButton_Name.Click += delegate (object o, EventArgs e)
            {
                blackWhiteList_Name = !blackWhiteList_Name;
                blackWhiteListButton_Name.Text = blackWhiteList_Name ? "Name Black List" : "Name White List";
            };

            TextBox blackWhiteListTextBox_Name = new TextBox
            {
                AutoSize = true,
                Margin = new Padding(6, 6, 6, 6),
                Text = blackWhiteListString_Name,
                Dock = DockStyle.Fill
            };
            l.Controls.Add(blackWhiteListTextBox_Name);

            tt.SetToolTip(blackWhiteListTextBox_Name, "Regular expression. Ignore Case. Use \";\" to separate. For exsample: \"Color Grading;Stabiliz[e(ation)]\".\nIf you want to modify the default value, you can edit the \".cs\" script file.\nIf you need to check the Video FX list, please refer to: https://zzzzzz9125.github.io/VegTips/videofxlist\nClick the TextBox with the mouse middle button to navigate to this webpage.");

            blackWhiteListTextBox_Name.MouseUp += delegate (object o, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Middle)
                {
                    System.Diagnostics.Process.Start("https://zzzzzz9125.github.io/VegTips/videofxlist");
                }
            };

            Button blackWhiteListButton_UID = new Button
            {
                Text = blackWhiteList_UID ? "UID Black List" : "UID White List",
                Margin = new Padding(0, 3, 0, 3),
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                Anchor = AnchorStyles.None
            };
            l.Controls.Add(blackWhiteListButton_UID);
            blackWhiteListButton_UID.FlatAppearance.BorderSize = 0;

            tt.SetToolTip(blackWhiteListButton_UID, "Click to toggle between UID Black List and UID White List.");

            blackWhiteListButton_UID.Click += delegate (object o, EventArgs e)
            {
                blackWhiteList_UID = !blackWhiteList_UID;
                blackWhiteListButton_UID.Text = blackWhiteList_UID ? "UID Black List" : "UID White List";
            };

            TextBox blackWhiteListTextBox_UID = new TextBox
            {
                AutoSize = true,
                Margin = new Padding(6, 6, 6, 6),
                Text = blackWhiteListString_UID,
                Dock = DockStyle.Fill
            };
            l.Controls.Add(blackWhiteListTextBox_UID);

            tt.SetToolTip(blackWhiteListTextBox_UID, "Regular expression. Ignore Case. Use \";\" to separate. For exsample: \"colorgrading;stabiliz[e(ation)]\".\nIf you want to modify the default value, you can edit the \".cs\" script file.\nIf you need to check the Video FX list, please refer to: https://zzzzzz9125.github.io/VegTips/videofxlist\nClick the TextBox with the mouse middle button to navigate to this webpage.");

            blackWhiteListTextBox_UID.MouseUp += delegate (object o, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Middle)
                {
                    System.Diagnostics.Process.Start("https://zzzzzz9125.github.io/VegTips/videofxlist");
                }
            };

            CheckBox randomPresetsCheckBox = new CheckBox
            {
                Text = "Random Presets",
                Margin = new Padding(6, 6, 6, 6),
                AutoSize = true,
                Checked = randomPresets
            };
            l.Controls.Add(randomPresetsCheckBox);
            l.SetColumnSpan(randomPresetsCheckBox, 2);

            tt.SetToolTip(randomPresetsCheckBox, "Toggle Random Presets");

            FlowLayoutPanel panel = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Anchor = AnchorStyles.None,
                Font = new Font("Arial", 8)
            };
            l.Controls.Add(panel);
            l.SetColumnSpan(panel, 2);

            Button ok = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK
            };
            panel.Controls.Add(ok);
            form.AcceptButton = ok;

            Button cancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel
            };
            panel.Controls.Add(cancel);
            form.CancelButton = cancel;

            DialogResult result = form.ShowDialog(myVegas.MainWindow);
            int temp;
            FxCount = int.TryParse(fxCountTextBox.Text, out temp) && temp > 0 ? temp : 1;
            objectChoice = cb.SelectedIndex;
            blackWhiteListString_Name = blackWhiteListTextBox_Name.Text;
            blackWhiteListString_UID = blackWhiteListTextBox_UID.Text;
            randomPresets = randomPresetsCheckBox.Checked;
            return result == DialogResult.OK;
        }
    }
}

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    //public void FromVegas(Vegas vegas, String scriptFile, XmlDocument scriptSettings, ScriptArgs args)
    {
        AddRandomFX.AddRandomFXClass test = new AddRandomFX.AddRandomFXClass();
        test.Main(vegas);
    }
}