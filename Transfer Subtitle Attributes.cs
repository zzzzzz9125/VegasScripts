using System;
using System.Collections;
using System.Windows.Forms;
using System.Collections.Generic;

using ScriptPortal.Vegas;  // "ScriptPortal.Vegas" for Magix Vegas Pro 14 or above, "Sony.Vegas" for Sony Vegas Pro 13 or below

namespace Test_Script
{
    public class Class
    {
        public Vegas myVegas;
        public void Main(Vegas vegas)
        {
            myVegas = vegas;
            myVegas.ResumePlaybackOnScriptExit = true;
            Project project = myVegas.Project;
            if (vegas.Project.Tracks.Count == 0)
            {
                return;
            }

            List<VideoEvent> selectList = new List<VideoEvent>();

            foreach (Track myTrack in project.Tracks)
            {
                if (myTrack.IsVideo())
                {
                    foreach (TrackEvent evnt in myTrack.Events)
                    {
                        if (evnt.Selected)
                        {
                            VideoEvent vEvent = (VideoEvent) evnt;
                            if (vEvent.ActiveTake.Media.IsGenerated())
                            {
                                selectList.Add(vEvent);
                            }
                        }
                    }
                }
            }

            foreach (VideoEvent vEvent in selectList)
            {
                // vEvent.OpenMediaGeneratorUI();
                Media sourceMedia = vEvent.ActiveTake.Media;
                OFXEffect sourceOFX = GetOFXEffect(sourceMedia);
                if (sourceOFX == null)
                {
                    continue;
                }

                bool isTAT = sourceMedia.Generator.PlugIn.UniqueID.ToLower().Contains("titlesandtext");

                OFXStringParameter textParam = FindTextParameter(sourceOFX);

                foreach (VideoEvent targetEvent in vEvent.Track.Events)
                {
                    if (targetEvent == vEvent)
                    {
                        continue;
                    }

                    Media targetMedia = targetEvent.ActiveTake.Media;
                    OFXEffect targetOFX = GetOFXEffect(targetMedia);
                    if (targetOFX == null || targetOFX.Label != sourceOFX.Label)
                    {
                        continue;
                    }

                    List<string> excludedParaNames = new List<string>();

                    OFXStringParameter targetTextPara = FindTextParameter(targetOFX);
                    if (textParam != null && targetTextPara != null)
                    {
                        string txt = targetTextPara.Value;

                        if (isTAT)
                        {
                            RichTextBox rtb = new RichTextBox();
                            rtb.Rtf = txt;
                            txt = rtb.Text;
                            rtb.Rtf = textParam.Value;
                            rtb.Text = txt;
                            txt = rtb.Rtf;
                        }

                        targetTextPara.Value = txt;
                        excludedParaNames.Add(targetTextPara.Name);
                    }

                    VideoStream sourceStream = (VideoStream) sourceMedia.Streams[0], targetStream = (VideoStream) targetMedia.Streams[0];
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
        
        public static OFXEffect GetOFXEffect(Media media)
        {
            if (media == null)
            {
                return null;
            }

            Effect generator = media.Generator;
            if (generator == null)
            {
                return null;
            }

            OFXEffect sourceOFX = generator.OFXEffect;
            if (sourceOFX == null)
            {
                return null;
            }

            return sourceOFX;
        }

        public static OFXStringParameter FindTextParameter(OFXEffect ofx)
        {
            // Compatible with other Text plugins, such as TextOFX (https://text.openfx.no/) and Textuler (https://textuler.io/)
            foreach (OFXParameter para in ofx.Parameters)
            {
                if (para.ParameterType == OFXParameterType.String && (para.Name.ToLower() == "text" || para.Name.ToLower() == "texts" || para.Name.ToLower() == "textmessage"))
                {
                    return (OFXStringParameter)para;
                }
            }
            return null;
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