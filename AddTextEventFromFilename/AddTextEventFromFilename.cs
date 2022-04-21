/*   Vegas Pro 19.0 用のスクリプト
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
using System.Text;
using System.Windows.Forms;
using ScriptPortal.Vegas;

using System.Drawing;



namespace vegastest1
{

    public class EntryPoint
    {
        public void FromVegas(Vegas vegas)
        {
            if (vegas.Project.Tracks.Count == 0)
            {
                // トラックがないときは何もしない
                return;
            }

            // パラメータを設定するダイアログを開く
            Tuple<DialogResult, string, string, string> dialogResult = DoDialog(vegas);
            DialogResult result = dialogResult.Item1;
            if (result != DialogResult.OK)
            {
                return;
            }

            string textFadeLengthString = dialogResult.Item2;
            string textLengthString = dialogResult.Item3;
            string textFadeString = dialogResult.Item4;

            // ダイアログを開き出力するログのファイルパスをユーザーに選択させる
            string saveFilePath = GetFilePath(vegas.Project.FilePath, "AddTextEventFromFilename");
            if (saveFilePath.Length == 0)
            {
                return;
            }

            // タイムライン上にある Video トラックの静止画・動画のファイル名を全部集める
            List<Tuple<long, string>> fileInfos = new List<Tuple<long, string>>();
            foreach (Track track in vegas.Project.Tracks)
            {
                string oldComment = "";

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

                    string filepath = GetMediaFilePath(take.Media);
                    if (filepath.Length == 0)
                    {
                        // ファイルパスが正常に取れなければ無視
                        continue;
                    }

                    // ファイルパスが見つかった

                    // 前回のコメントと同じなら無視する
                    string comment = GetComment(filepath); // 今回のコメント
                    if (comment == oldComment)
                    {
                        // 前回と同じコメントなので無視する
                        continue;
                    }

                    oldComment = comment;

                    // このファイルが張り付けられているフレーム位置とファイルパスのペアを追加する
                    fileInfos.Add(Tuple.Create(trackEvent.Start.FrameCount, filepath));                    
                }
            }

            // Titles & Text の Generator を取得する
            PlugInNode generator = GetGeneratorTitlesAndText(vegas);
            if (generator == null)
            {
                MessageBox.Show("Cannot get Titles & Text generator");
                return;
            }

            // ログ出力用のファイルを開く
            System.IO.StreamWriter writer = new System.IO.StreamWriter(saveFilePath, false, Encoding.GetEncoding("Shift_JIS"));


            // 抽出したファイル名を、最初のトラックの同時刻に追加する
            Timecode textFadeLength = Timecode.FromString(textFadeLengthString);   // 追加するテキストのフェード時間
            VideoTrack textVideoTrack = vegas.Project.Tracks[0] as VideoTrack;    // テキストを追加するビデオトラック（先頭のトラックに追加する）
            if (textVideoTrack == null)
            {
                return;
            }

            long textEventOffsetMs = 1000;  // 追加するテキストは、動画や静止画よりも少し遅らせて追加する。遅らせる時間を ms で指定する
            if (!long.TryParse(textFadeString, out textEventOffsetMs))
            {
                MessageBox.Show("Cannot parse Text Offset");
                return;
            }

            Timecode textLength = Timecode.FromString(textLengthString);  // テキストの表示時間。固定時間としているが、動画や静止画がこれより短いとテキストが正しく表示されないので、動画や静止画の長さをチェックする必要がある

            // ファイル名からテキストイベントを作成する
            foreach (var fileInfo in fileInfos)
            {
                string filepath = fileInfo.Item2;
                Timecode timecodeStart = new Timecode();
                timecodeStart.FrameCount = fileInfo.Item1;

                // TextEvent を表す新しい Media を生成する
                Media media = Media.CreateInstance(vegas.Project, generator);

                // Media の Effect を取得する
                OFXEffect ofxEffect = GetOFXEffect(media);
                if (ofxEffect == null)
                {
                    continue;
                }

                // 取得した Effect のパラメータを変更して、望むテロップにする

                long err = 0;

                // テキストを変える
                {
                    string newText = ToComment(filepath); // ファイル名からコメントに変換する
                    float fontSize = 14;
                    err = ChangeText(ofxEffect, newText, fontSize);
                    if (err != 0)
                    {
                        // エラーは無視
                    }
                }

                // テキストの表示位置を変える
                {
                    double x = 0.5;
                    double y = 0.1;
                    err = ChangeLocation(ofxEffect, x, y);
                    if (err != 0)
                    {
                        // エラーは無視
                    }
                }

                // アウトラインの幅を変える
                {
                    double outlineWidth = 10.0;
                    err = ChangeOutlineWidth(ofxEffect, outlineWidth);
                    if (err != 0)
                    {
                        // エラーは無視
                    }
                }

                // アウトラインの色を変える
                {
                    double r = 0.0;
                    double g = 0.0;
                    double b = 0.0;
                    double a = 1.0;
                    err = ChangeOutlineColor(ofxEffect, r, g, b, a);
                    if (err != 0)
                    {
                        // エラーは無視
                    }
                }

                ofxEffect.AllParametersChanged();

                // テロップを入れたい位置に video Event を生成する
                Timecode textStartTimecode = timecodeStart + Timecode.FromMilliseconds(textEventOffsetMs);
                VideoEvent videoEvent = new VideoEvent(textStartTimecode, textLength);

                // テキストをトラックに追加する
                textVideoTrack.Events.Add(videoEvent);

                Take take = new Take(media.GetVideoStreamByIndex(0));
                videoEvent.Takes.Add(take);

                videoEvent.FadeIn.Length = textFadeLength;
                videoEvent.FadeOut.Length = textFadeLength;

                // ログに出力する
                writer.WriteLine(timecodeStart.ToString() + " " + ToComment(filepath).Replace('\n', ' '));
            }

            writer.Close();


            MessageBox.Show("終了しました。");
        }


        private Tuple<DialogResult, string, string, string> DoDialog(Vegas vegas)
        {
            Form form = new Form();
            form.SuspendLayout();
            form.AutoScaleMode = AutoScaleMode.Font;
            form.AutoScaleDimensions = new SizeF(6F, 13F);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterParent;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.HelpButton = false;
            form.ShowInTaskbar = false;
            form.Text = "Add Text Event From Filename";
            form.AutoSize = true;
            form.AutoSizeMode = AutoSizeMode.GrowAndShrink;

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.AutoSize = true;
            layout.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            layout.GrowStyle = TableLayoutPanelGrowStyle.AddRows;
            layout.ColumnCount = 3;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            form.Controls.Add(layout);

            {
                Label label;

                label = new Label();
                label.Text = "Text Fade Length:";
                label.AutoSize = false;
                label.TextAlign = ContentAlignment.MiddleLeft;
                label.Margin = new Padding(8, 8, 8, 4);
                label.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                layout.Controls.Add(label);
                layout.SetColumnSpan(label, 3);
            }

            TextBox TextLFadeLengthBox = new TextBox();
            {
                TextLFadeLengthBox.Text = "00:00:01;00";

                TextLFadeLengthBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                TextLFadeLengthBox.Margin = new Padding(16, 8, 8, 4);
                layout.Controls.Add(TextLFadeLengthBox);
                layout.SetColumnSpan(TextLFadeLengthBox, 2);
            }

            {
                Label label;

                label = new Label();
                label.Text = "Text Length:";
                label.AutoSize = false;
                label.TextAlign = ContentAlignment.MiddleLeft;
                label.Margin = new Padding(8, 8, 8, 4);
                label.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                layout.Controls.Add(label);
                layout.SetColumnSpan(label, 3);
            }

            TextBox TextLengthBox = new TextBox();
            {
                TextLengthBox.Text = "00:00:05;00";

                TextLengthBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                TextLengthBox.Margin = new Padding(16, 8, 8, 4);
                layout.Controls.Add(TextLengthBox);
                layout.SetColumnSpan(TextLengthBox, 2);
            }

            {
                Label label;

                label = new Label();
                label.Text = "Text Offset (ms):";
                label.AutoSize = false;
                label.TextAlign = ContentAlignment.MiddleLeft;
                label.Margin = new Padding(8, 8, 8, 4);
                label.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                layout.Controls.Add(label);
                layout.SetColumnSpan(label, 3);
            }

            TextBox TextOffsetBox = new TextBox();
            {
                TextOffsetBox.Text = "1000";

                TextOffsetBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                TextOffsetBox.Margin = new Padding(16, 8, 8, 4);
                layout.Controls.Add(TextOffsetBox);
                layout.SetColumnSpan(TextOffsetBox, 2);
            }

            {
                FlowLayoutPanel buttonPanel = new FlowLayoutPanel();
                buttonPanel.FlowDirection = FlowDirection.RightToLeft;
                buttonPanel.Size = Size.Empty;
                buttonPanel.AutoSize = true;
                buttonPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                buttonPanel.Margin = new Padding(8, 8, 8, 8);
                buttonPanel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                layout.Controls.Add(buttonPanel);
                layout.SetColumnSpan(buttonPanel, 3);

                {
                    Button cancelButton = new Button();
                    cancelButton.Text = "Cancel";
                    cancelButton.FlatStyle = FlatStyle.System;
                    cancelButton.DialogResult = DialogResult.Cancel;
                    buttonPanel.Controls.Add(cancelButton);
                    form.CancelButton = cancelButton;
                }

                {
                    Button okButton = new Button();

                    okButton.Text = "OK";
                    okButton.FlatStyle = FlatStyle.System;
                    okButton.DialogResult = DialogResult.OK;
                    buttonPanel.Controls.Add(okButton);
                    form.AcceptButton = okButton;
                }
            }

            form.ResumeLayout();

            string textFadeLength = "";
            string textLength = "";
            string textOffset = "";
            DialogResult result = form.ShowDialog(vegas.MainWindow);
            if (DialogResult.OK == result)
            {
                textFadeLength = TextLFadeLengthBox.Text;
                textLength = TextLengthBox.Text;
                textOffset = TextOffsetBox.Text;
            }

            return Tuple.Create(result, textFadeLength, textLength, textOffset);
        }

        // Titles & Text のジェネレータを取得する
        private PlugInNode GetGeneratorTitlesAndText(Vegas vegas)
        {
            PlugInNode generator = vegas.Generators.GetChildByUniqueID("{Svfx:com.vegascreativesoftware:titlesandtext}"); //  Titles & Text
            return generator;
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

        private OFXEffect GetOFXEffect(Media media)
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

            OFXEffect ofxEffect = generator.OFXEffect;
            if (ofxEffect == null)
            {
                return null;
            }

            return ofxEffect;
        }

        private string GetMediaFilePath(Media media)
        {
            if (media == null)
            {
                return "";
            }

            string filepath = media.FilePath;
            if (filepath.Length == 0)
            {
                // ファイルパスが正常に取れなかった
                return "";
            }

            // テキストイベントにもテキストの内容がファイルパスに入っている
            // それはほしいものではないので（そこにファイルはないので）除外する
            if (IsTextEvent(media))
            {
                // テキストイベントは無視
                return "";
            }

            return filepath;
        }

        // media がテキストイベントであれば true
        private bool IsTextEvent(Media media)
        {
            OFXEffect ofxEffect = GetOFXEffect(media);
            if (ofxEffect == null)
            {
                return false;
            }

            OFXStringParameter textParam = ofxEffect.FindParameterByName("Text") as OFXStringParameter;
            if (textParam == null)
            {
                // テキストイベントではない
                return false;
            }

            return true;
        }


        // filepathから撮影日を取得する
        private string GetDate(string filepath)
        {
            string filename = Path.GetFileNameWithoutExtension(filepath);

            // filenameから撮影日を取得する
            var firstTIndex = filename.IndexOf('T');
            if (firstTIndex < 0)
            {
                return "";
            }

            string date_ = filename.Substring(0, firstTIndex);
            string date = date_.Replace('-', '.');

            return date;
        }

        // filepathからコメントを取得する
        private string GetComment(string filepath)
        {
            string filename = Path.GetFileNameWithoutExtension(filepath);

            // filenameからコメントを取得する
            var firstUnderbarIndex = filename.IndexOf('_');
            if (firstUnderbarIndex < 0)
            {
                return "";
            }

            string comment = filename.Substring(firstUnderbarIndex + 1);

            return comment;
        }

        // ファイル名からテキストイベントのテキストに変換する
        // ファイル名は  2021-06-30T145043_comment.mp4 のようになっていることが前提
        // 以下のようなテキストが得られる：
        // 2021.06.30
        // comment
        private  string ToComment(string filepath)
        {
            string date = GetDate(filepath);
            string comment = GetComment(filepath);

            // 撮影日とコメントを連結する
            string text = date + " \n" + comment;

            return text;
        }

        // テキストを変える
        private long ChangeText(OFXEffect ofxEffect, string newText, float fontSize)
        {
            OFXStringParameter textParam = ofxEffect.FindParameterByName("Text") as OFXStringParameter;
            if (textParam == null)
            {
                return -1;
            }

            {
                string rtfData = textParam.Value;   // デフォルトで入っているテキスト

                RichTextBox richtextBox = new RichTextBox();
                FontFamily fontFamily = richtextBox.SelectionFont.FontFamily;
                richtextBox.Rtf = rtfData;
                richtextBox.Text = newText;
                richtextBox.SelectAll();    // 全テキストが対象

                richtextBox.SelectionFont = new System.Drawing.Font(fontFamily, fontSize);    // フォント変更

                textParam.Value = richtextBox.Rtf;
            }

            return 0;
        }

        // テキストの表示位置を変える
        private long ChangeLocation(OFXEffect ofxEffect, double x, double y)
        {
            OFXDouble2DParameter locationParam = ofxEffect.FindParameterByName("Location") as OFXDouble2DParameter;
            if (locationParam == null)
            {
                return -1;
            }

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
            if (outlineWidthParam == null)
            {
                return -1;
            }

            outlineWidthParam.Value = width;

            return 0;
        }

        // アウトラインの色を変える
        private long ChangeOutlineColor(OFXEffect ofxEffect, double r, double g, double b, double a)
        {
            OFXRGBAParameter outlineColorParam = ofxEffect.FindParameterByName("OutlineColor") as OFXRGBAParameter;
            if (outlineColorParam == null)
            {
                return -1;
            }

            outlineColorParam.Value = new OFXColor(r, g, b, a);
  
            return 0;
        }

    }

}
