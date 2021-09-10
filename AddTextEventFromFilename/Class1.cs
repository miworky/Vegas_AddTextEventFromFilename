﻿/*   Vegas Pro 18.0 用のプラグイン
 * 　　タイムライン上にある Video トラックの静止画・動画のファイル名を元に、テロップを自動生成する。
 * 　　静止画・動画のファイル名は「YYYY-MM-DDTHHmmSS_コメント」となっている必要がある。
 * 　　
 * 　　miteneDownloaderを使用してみてねからダウンロードしたファイルをタイムラインに配置していることを想定している。
 * 　　
 * 　　miteneDownloader：
 * 　　https://github.com/miworky/miteneDownloader
 */

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ScriptPortal.Vegas;

using System.Drawing;



namespace vegastest1
{
    static class VegasConstants
    {
        public const string TextEventFileNamePrefix = "VEGAS タイトルおよびテキスト";
    }


    public class EntryPoint
    {
        public void FromVegas(Vegas vegas)
        {
            if (vegas.Project.Tracks.Count == 0)
            {
                // トラックがないときは何もしない
                return;
            }

            // ダイアログを開き出力するログのファイルパスをユーザーに選択させる
            string saveFilePath = GetFilePath(vegas.Project.FilePath, "AddTextEventFromFilename");

            // タイムライン上にある Video トラックの静止画・動画のファイル名を全部集める
            List<Tuple<long, string>> fileNames = new List<Tuple<long, string>>();
            foreach (Track track in vegas.Project.Tracks)
            {
                foreach (TrackEvent trackEvent in track.Events)
                {
                    if (!trackEvent.IsVideo())
                    {
                        // ビデオトラック以外は無視する
                        continue;
                    }

                    // アクティブテイクのみを対象にする
                    Take take = trackEvent.ActiveTake;
                    if (take == null)
                    {
                        // アクティブテイクがなければ無視
                        continue;
                    }

                    string filepath = take.Media.FilePath;
                    if (filepath.Length == 0)
                    {
                        // ファイルパスが正常に取れなければ無視
                        continue;
                    }

                    if (filepath.StartsWith(VegasConstants.TextEventFileNamePrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        // テキストイベントは無視
                        continue;
                    }

                    // ファイルパスが見つかった
                    // このファイルが張り付けられているフレーム位置とファイルパスのペアを追加する
                    fileNames.Add(Tuple.Create(trackEvent.Start.FrameCount, filepath));                    
                }
            }


            // ログ出力用のファイルを開く
            System.IO.StreamWriter writer = new System.IO.StreamWriter(saveFilePath, false, Encoding.GetEncoding("Shift_JIS"));


            // 抽出したファイル名を、最初のトラックの同時刻に追加する
            Timecode textFadeLength = Timecode.FromString("00:00:01;00");   // 追加するテキストのフェード時間
            VideoTrack textVideoTrack = (VideoTrack)vegas.Project.Tracks[0];    // テキストを追加するビデオトラック（先頭のトラックに追加する）
            long textEventOffsetMs = 1000;  // 追加するテキストは、動画や静止画よりも少し遅らせて追加する。遅らせる時間を ms で指定する
            Timecode textLength = Timecode.FromString("00:00:05;00");  // テキストの表示時間。固定時間としているが、動画や静止画がこれより短いとテキストが正しく表示されないので、動画や静止画の長さをチェックする必要がある

            foreach (var frameCountFilename in fileNames)
            {
                string name = frameCountFilename.Item2;
                Timecode timecode = new Timecode();
                timecode.FrameCount = frameCountFilename.Item1;

                PlugInNode generator = vegas.Generators.GetChildByName("VEGAS Titles & Text");
                Media media = Media.CreateInstance(vegas.Project, generator);

                OFXEffect ofxEffect = media.Generator.OFXEffect;

                // テキストを変える
                {
                    string newText = ToComment(name); // ファイル名からコメントに変換する
                    float fontSize = 14;
                    ChangeText(ofxEffect, newText, fontSize);
                }

                // テキストの表示位置を変える
                {
                    double x = 0.5;
                    double y = 0.1;
                    ChangeLocation(ofxEffect, x, y);
                }

                // アウトラインの幅を変える
                {
                    double outlineWidth = 10.0;
                    ChangeOutlineWidth(ofxEffect, outlineWidth);
                }

                // アウトラインの色を変える
                {
                    double r = 0.0;
                    double g = 0.0;
                    double b = 0.0;
                    double a = 1.0;
                    ChangeOutlineColor(ofxEffect, r, g, b, a);
                }

                ofxEffect.AllParametersChanged();


                Timecode textStartTimecode = timecode + Timecode.FromMilliseconds(textEventOffsetMs);

                VideoEvent videoEvent = new VideoEvent(textStartTimecode, textLength);

                // テキストをトラックに追加する
                textVideoTrack.Events.Add(videoEvent);

                Take take = new Take(media.GetVideoStreamByIndex(0));
                videoEvent.Takes.Add(take);

                videoEvent.FadeIn.Length = textFadeLength;
                videoEvent.FadeOut.Length = textFadeLength;

                // ログに出力する
                writer.WriteLine(timecode.ToString() + " " + ToComment(name).Replace('\n', ' '));
            }

            writer.Close();


            MessageBox.Show("終了しました。");
        }


        // ダイアログを開きファイルパスをユーザーに選択させる
        private string GetFilePath(string rootFilePath, string preFix)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = preFix + System.IO.Path.GetFileNameWithoutExtension(rootFilePath) + ".txt";
            sfd.InitialDirectory = System.IO.Path.GetDirectoryName(rootFilePath) + "\\";
            sfd.Filter = "テキストファイル(*.txt)|*.txt";
            if (sfd.ShowDialog() != DialogResult.OK)
            {
                return "";
            }

            return sfd.FileName;
        }

        // ファイル名からテキストイベントのテキストに変換する
        // ファイル名は  2021-06-30T145043_comment.mp4 のようになっていることが前提
        // 以下のようなテキストが得られる：
        // 2021.06.30
        // comment

       private  string ToComment(string name)
        {
            string filename = Path.GetFileNameWithoutExtension(name);

            // filenameから撮影日を取得する
            var firstTIndex = filename.IndexOf('T');
            if (firstTIndex < 0)
            {
                return "";
            }

            string date_ = filename.Substring(0, firstTIndex);
            string date = date_.Replace('-', '.');

            // filenameからコメントを取得する
            var firstUnderbarIndex = filename.IndexOf('_');
            if (firstUnderbarIndex < 0)
            {
                return "";
            }

            string comment = filename.Substring(firstUnderbarIndex + 1);

            // 撮影日とコメントを連結する
            string text = date + " \n" + comment;

            return text;
        }

        // テキストを変える
        private long ChangeText(OFXEffect ofxEffect, string newText, float fontSize)
        {
            OFXStringParameter textParam = ofxEffect.FindParameterByName("Text") as OFXStringParameter;
            {
                string rtfData = textParam.Value;   // デフォルトで入っているテキスト

                RichTextBox richtextBox = new RichTextBox();
                richtextBox.Rtf = rtfData;
                richtextBox.Text = newText;
                richtextBox.SelectAll();    // 全テキストが対象

                FontFamily fontFamily = richtextBox.SelectionFont.FontFamily;
                richtextBox.SelectionFont = new System.Drawing.Font(fontFamily, fontSize);    // フォント変更

                textParam.Value = richtextBox.Rtf;
            }

            return 0;
        }

        // テキストの表示位置を変える
        private long ChangeLocation(OFXEffect ofxEffect, double x, double y)
        {
            OFXDouble2DParameter locationParam = ofxEffect.FindParameterByName("Location") as OFXDouble2DParameter;

            OFXDouble2D location;
            location.X = x;
            location.Y = y;

            locationParam.Value = location;

            return 0;
        }

        // アウトラインの幅を変える
        private long ChangeOutlineWidth(OFXEffect ofxEffect, double width)
        {
            OFXDoubleParameter outlineWidthParam = ofxEffect.FindParameterByName("OutlineWidth") as OFXDoubleParameter;
            outlineWidthParam.Value = width;

            return 0;
        }

        // アウトラインの色を変える
        private long ChangeOutlineColor(OFXEffect ofxEffect, double r, double g, double b, double a)
        {
            OFXRGBAParameter outlineColorParam = ofxEffect.FindParameterByName("OutlineColor") as OFXRGBAParameter;
            outlineColorParam.Value = new OFXColor(r, g, b, a);
  
            return 0;
        }

    }

}