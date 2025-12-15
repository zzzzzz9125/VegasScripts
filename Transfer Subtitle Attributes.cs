using System;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using ScriptPortal.Vegas;  // "ScriptPortal.Vegas" for Magix Vegas Pro 14 or above, "Sony.Vegas" for Sony Vegas Pro 13 or below

namespace Test_Script
{
    public class Class
    {
        public const bool DO_POP_UP_WINDOW = true;
        public const int ITEMS_PER_LINE = 5;
        public Vegas myVegas;
        public void Main(Vegas vegas)
        {
            myVegas = vegas;
            myVegas.ResumePlaybackOnScriptExit = true;
            if (myVegas.Project.Tracks.Count == 0)
            {
                return;
            }

            List<VideoEvent> evs = new List<VideoEvent>();

            foreach (Track myTrack in myVegas.Project.Tracks)
            {
                if (myTrack.IsVideo())
                {
                    foreach (TrackEvent ev in myTrack.Events)
                    {
                        if (ev.Selected)
                        {
                            if (ev.ActiveTake != null && ev.ActiveTake.Media != null && ev.ActiveTake.Media.IsGenerated() && ev.ActiveTake.Media.Generator.PlugIn.IsOFX)
                            {
                                evs.Add(ev as VideoEvent);
                            }
                        }
                    }
                }
            }

            Dictionary<string, List<string>> dicExclude = new Dictionary<string, List<string>>();

            foreach (VideoEvent ev in evs)
            {
                Effect ef = ev.ActiveTake.Media.Generator;
                if (dicExclude.ContainsKey(ef.PlugIn.UniqueID))
                {
                    continue;
                }
                OFXEffect ofx = ef.OFXEffect;
                List<string> exclude = new List<string>();
                dicExclude.Add(ef.PlugIn.UniqueID, exclude);
                if (!DO_POP_UP_WINDOW)
                {
                    continue;
                }
                Dictionary<string, string> dicParameters = new Dictionary<string, string>();
                
                foreach (OFXParameter para in ofx.Parameters)
                {
                    string type = para.ParameterType.ToString();
                    if (type == "Group" || type == "PushButton" || type == "16" || type == "Page")  // "16" is for "OfxParamTypeImage", which is not supported
                    {
                        continue;
                    }
                    string label = para.Label;
                    string groupName = para.ParentName;
                    if (!string.IsNullOrEmpty(groupName))
                    {
                        OFXParameter group = ofx.FindParameterByName(groupName);
                        if (group != null)
                        {
                            label = string.Format("[{0}] {1}", group.Label, para.Label);
                        }
                    }
                    dicParameters.Add(para.Name, label);
                }
                
                if (!ShowWindow(ofx.Label, dicParameters, exclude))
                {
                    return;
                }
            }

            foreach (VideoEvent ev in evs)
            {
                Media sourceMedia = ev.ActiveTake.Media;
                Effect sourceEffect = sourceMedia.Generator;
                OFXEffect sourceOFX = sourceEffect.OFXEffect;

                bool isTAT = sourceEffect.PlugIn.UniqueID.ToLower().Contains("titlesandtext");

                List<OFXStringParameter> textParas = FindTextParameters(sourceEffect);

                foreach (VideoEvent targetEvent in ev.Track.Events)
                {
                    if (targetEvent == ev || targetEvent.ActiveTake == null || targetEvent.ActiveTake.Media == null || !targetEvent.ActiveTake.Media.IsGenerated())
                    {
                        continue;
                    }

                    Media targetMedia = targetEvent.ActiveTake.Media;
                    Effect targetEffect = targetMedia.Generator;
                    if (!targetEffect.PlugIn.IsOFX || sourceEffect.PlugIn.UniqueID != targetEffect.PlugIn.UniqueID)
                    {
                        continue;
                    }

                    OFXEffect targetOFX = targetEffect.OFXEffect;

                    List<string> excludedParaNames = new List<string>();

                    if (dicExclude.ContainsKey(sourceEffect.PlugIn.UniqueID))
                    {
                        excludedParaNames = new List<string>(dicExclude[sourceEffect.PlugIn.UniqueID]);
                    }

                    foreach (OFXStringParameter textPara in textParas)
                    {
                        if (excludedParaNames.Contains(textPara.Name))
                        {
                            continue;
                        }
                        OFXStringParameter targetTextPara = targetOFX.FindParameterByName(textPara.Name) as OFXStringParameter;
                        if (targetTextPara == null)
                        {
                            continue;
                        }

                        string txt = targetTextPara.Value;

                        if (isTAT)
                        {
                            RichTextBox rtb = new RichTextBox();
                            rtb.Rtf = txt;
                            txt = rtb.Text;
                            rtb.Rtf = textPara.Value;
                            rtb.Text = txt;
                            txt = rtb.Rtf;
                        }

                        targetTextPara.Value = txt;
                        excludedParaNames.Add(targetTextPara.Name);
                    }


                    VideoStream sourceStream = sourceMedia.GetVideoStreamByIndex(0), targetStream = targetMedia.GetVideoStreamByIndex(0);
                    targetStream.Size = sourceStream.Size;
                    targetStream.FieldOrder = sourceStream.FieldOrder;
                    targetStream.PixelAspectRatio = sourceStream.PixelAspectRatio;
                    targetStream.AlphaChannel = sourceStream.AlphaChannel;
                    targetStream.BackgroundColor = sourceStream.BackgroundColor;

                    double adjustValue = 1.0 * targetStream.Length.Nanos / sourceMedia.Length.Nanos;

                    CopyOFX(sourceOFX, targetOFX, true, adjustValue, excludedParaNames);
                }
            }
        }

        public bool ShowWindow(string plugInName, Dictionary<string, string> dicParameters, List<string> exclude)
        {
            Form form = new Form
            {
                ShowInTaskbar = false,
                AutoSize = true,
                BackColor = Color.FromArgb(45, 45, 45),
                ForeColor = Color.FromArgb(200, 200, 200),
                Font = new Font("Arial", 9),
                Text = "Transfer Subtitle Attributes",
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
                ColumnCount = ITEMS_PER_LINE
            };
            p.Controls.Add(l);

            Label label = new Label
            {
                Margin = new Padding(12, 12, 6, 6),
                Text = string.Format("PlugIn Name: {0}", plugInName),
                AutoSize = true
            };
            l.Controls.Add(label);
            l.SetColumnSpan(label, l.ColumnCount);

            List<CheckBox> ckbs = new List<CheckBox>();
            foreach (string paraName in dicParameters.Keys)
            {
                CheckBox ckb = new CheckBox
                {
                    Text = dicParameters[paraName],
                    Margin = new Padding(6, 6, 6, 6),
                    AutoSize = true,
                    Checked = true,
                    Tag = paraName
                };
                l.Controls.Add(ckb);
                ckbs.Add(ckb);
            }

            FlowLayoutPanel panel = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Anchor = AnchorStyles.None,
                Font = new Font("Arial", 8)
            };
            l.Controls.Add(panel);
            l.SetColumnSpan(panel, l.ColumnCount);

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

            foreach (CheckBox ckb in ckbs)
            {
                if (!ckb.Checked)
                {
                    exclude.Add(ckb.Tag as string);
                }
            }

            return result == DialogResult.OK;
        }

        public static List<OFXStringParameter> FindTextParameters(Effect ef)
        {
            List<OFXStringParameter> paras = new List<OFXStringParameter>();
            if (!ef.PlugIn.IsOFX)
            {
                return paras;
            }
            // Compatible with other Text plugins, such as TextOFX (https://text.openfx.no/), Textuler (https://textuler.io/), OFX Clock (https://www.hlinke.de/dokuwiki/doku.php?id=en:vegas_pro_ofx) and Universe Typographic
            foreach (OFXParameter para in ef.OFXEffect.Parameters)
            {
                if (para.ParameterType.ToString() != "String" || string.IsNullOrEmpty(para.Name))
                {
                    continue;
                }

                bool is_Universe_Text_Typographic_OFX = ef.PlugIn.UniqueID.Contains("Universe_Text_Typographic_OFX");
                switch (para.Name.ToLower())
                {
                    case "text":
                    case "texts":
                    case "textmessage": // Textuler
                    case "Dig Clock Free Format": // OFX Clock
                        paras.Add(para as OFXStringParameter);
                        break;

                    case "68": // Universe Typographic
                    case "116": // Universe Typographic
                        if (is_Universe_Text_Typographic_OFX)
                        {
                            paras.Add(para as OFXStringParameter);
                        }
                        break;

                    default:
                        break;
                }
            }
            return paras;
        }

        public static void CopyOFX(OFXEffect sourceOFX, OFXEffect targetOFX, bool isMediaGenerator = false, double adjust = 0, List<string> exclude = null)
        {
            adjust = adjust > 0 ? adjust : 1;
            foreach (OFXParameter targetPara in targetOFX.Parameters)
            {
                OFXParameter para = sourceOFX.FindParameterByName(targetPara.Name);

                if (para == null || targetPara.ParameterType != para.ParameterType || (exclude != null && exclude.Contains(targetPara.Name)))
                {
                    continue;
                }

                switch (targetPara.ParameterType)
                {
                    case OFXParameterType.Boolean:
                    {
                        OFXBooleanParameter tp = (OFXBooleanParameter)targetPara, p = (OFXBooleanParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.Choice:
                    {
                        OFXChoiceParameter tp = (OFXChoiceParameter)targetPara, p = (OFXChoiceParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.Custom:
                    {
                        OFXCustomParameter tp = (OFXCustomParameter)targetPara, p = (OFXCustomParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.Double2D:
                    {
                        OFXDouble2DParameter tp = (OFXDouble2DParameter)targetPara, p = (OFXDouble2DParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.Double3D:
                    {
                        OFXDouble3DParameter tp = (OFXDouble3DParameter)targetPara, p = (OFXDouble3DParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.Double:
                    {
                        OFXDoubleParameter tp = (OFXDoubleParameter)targetPara, p = (OFXDoubleParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.Integer2D:
                    {
                        OFXInteger2DParameter tp = (OFXInteger2DParameter)targetPara, p = (OFXInteger2DParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.Integer3D:
                    {
                        OFXInteger3DParameter tp = (OFXInteger3DParameter)targetPara, p = (OFXInteger3DParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.Integer:
                    {
                        OFXIntegerParameter tp = (OFXIntegerParameter)targetPara, p = (OFXIntegerParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.RGBA:
                    {
                        OFXRGBAParameter tp = (OFXRGBAParameter)targetPara, p = (OFXRGBAParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.RGB:
                    {
                        OFXRGBParameter tp = (OFXRGBParameter)targetPara, p = (OFXRGBParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    case OFXParameterType.String:
                    {
                        OFXStringParameter tp = (OFXStringParameter)targetPara, p = (OFXStringParameter)para;
                        tp.IsAnimated = p.IsAnimated;
                        if (p.IsAnimated)
                        {
                            tp.Keyframes.Clear();
                            for (int i = 0; i < p.Keyframes.Count; i++)
                            {
                                Timecode tc = Timecode.FromNanos((long)(p.Keyframes[i].Time.Nanos * adjust * (isMediaGenerator ? (p.Keyframes[i].Time.FrameRate / 1000) : 1)));
                                if (i == 0)
                                {
                                    tp.Keyframes[i].Time = tc;
                                    tp.Keyframes[i].Value = p.Keyframes[i].Value;
                                }
                                tp.SetValueAtTime(tc, p.Keyframes[i].Value);
                                tp.Keyframes[i].Interpolation = p.Keyframes[i].Interpolation;
                            }
                        }
                        else { tp.Value = p.Value; } break;
                    }

                    default: break;
                }
            }

            targetOFX.AllParametersChanged();
        }
    }
}

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    //public void FromVegas(Vegas vegas, String scriptFile, XmlDocument scriptSettings, ScriptArgs args)
    {
        Test_Script.Class test = new Test_Script.Class();
        test.Main(vegas);
        Application.Exit();
    }
}