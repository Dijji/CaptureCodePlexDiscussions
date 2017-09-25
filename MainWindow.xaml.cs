using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using WatiN.Core;
using Microsoft.Office.Interop.Word;
using Find = WatiN.Core.Find;
using System.Threading;
using System.Xml.Serialization;

namespace GetDiscussions
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Get_Click(object sender, RoutedEventArgs e)
        {
            status.Content = "Reading from CodePlex...";
            var siteName = site.Text.Trim();
            var toWord = (bool)checkWord.IsChecked;
            var toXml = (bool)checkXml.IsChecked;

            var task = StartSTATask<string>(() =>
            {
                return GetDiscussions(siteName, toWord, toXml);
            })
            .ContinueWith((t) =>
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    if (t.Result == null)
                        status.Content = "Success!";
                    else
                        status.Content = "Get failed with: " + t.Result;
                });
            });
        }

        private string GetDiscussions(string siteName, bool toWord, bool toXml)
        {
            try
            {
                List<Discussion> discs = new List<Discussion>();
                using (var ie = new IE())
                {
                    var ds = GetDiscussionUrls(ie, String.Format("https://{0}.codeplex.com/", siteName));
                    foreach (var url in ds)
                    {
                        discs.Add(GetDiscussion(ie, url));
                    }
                }

                if (toWord)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        status.Content = "Creating Word document…";
                    });

                    var word = new Microsoft.Office.Interop.Word.Application();
                    object missing = System.Reflection.Missing.Value;
                    var doc = word.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                    var HtmlTempPath = Path.Combine(Path.GetTempPath(), $"{Path.GetRandomFileName()}.html");

                    foreach (var disc in discs)
                    {
                        Paragraph para2 = doc.Content.Paragraphs.Add(ref missing);
                        // Text must be set before Style for the style to be applied properly
                        para2.Range.Text = disc.Title;
                        para2.Range.set_Style(WdBuiltinStyle.wdStyleHeading2);
                        para2.Range.InsertParagraphAfter();

                        if (disc.Posts.Count > 0)
                        {
                            var table = doc.Tables.Add(para2.Range, disc.Posts.Count, 2, ref missing, ref missing);
                            table.AllowAutoFit = true;
                            table.Borders.Enable = 1;


                            foreach (Row row in table.Rows)
                            {
                                var p = disc.Posts[row.Index - 1];
                                row.Cells[1].Range.Text = p.From + "\n" + p.When;

                                File.WriteAllText(HtmlTempPath, $"<html>{p.Html}</html>");
                                row.Cells[2].Range.InsertFile(HtmlTempPath, ref missing, ref missing, ref missing, ref missing);
                            }
                        }
                    }
                    
                    doc.SaveAs2(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), siteName + " Discussions.docx"));
                    doc.Close();
                    word.Quit();
                    File.Delete(HtmlTempPath);
                }

                if (toXml)
                {
                    // Output XML
                    var ds = new DiscussionSet { Discussions = discs };
                    XmlSerializer x = new XmlSerializer(typeof(DiscussionSet));
                    using (TextWriter writer = new StreamWriter(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), siteName + " Discussions.xml")))
                    {
                        x.Serialize(writer, ds);
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private List<string> GetDiscussionUrls(IE ie, string siteUrl)
        {
            List<string> discussions = new List<string>();
            ie.GoTo(siteUrl);
            ie.WaitForComplete();
            
            bool done = false;
            ie.Link(Find.ById("discussionTab")).Click();
            while (!done)
            {
                ie.WaitForComplete();
                foreach (var div in ie.Divs)
                {
                    if (div.ClassName == "post_content")
                    {
                        discussions.Add(div.Link(Find.First()).Url);
                    }
                }

                var pag = ie.List(Find.ById("discussion_pagination"));
                var link = pag.Link(Find.ByText("Next"));
                if (link.Exists)
                    link.Click();
                else
                    done = true;
            }

            return discussions;
        }

        private Discussion GetDiscussion(IE ie, string url)
        {
            var disc = new Discussion();
            ie.GoTo(url);
            ie.WaitForComplete();

            var hDiv = ie.Div(Find.ByClass("ViewThread"));
            var h = hDiv.Elements.First();
            disc.Title = h.InnerHtml.Replace("\n", "").Trim();

            var tDiv = ie.Div(Find.ByClass("Posts"));
            foreach(var tr in tDiv.TableRows)
            {
                if (tr.Id == "PostPanel")
                {
                    var p = new Post();
                    var details = tr.TableCells[0];
                    p.From = details.Div(Find.ByClass("UserName")).Elements.First().InnerHtml;
                    p.When = details.Span(Find.ByClass("smartDate")).Title;
                    var content = tr.TableCells[1];
                    p.Html = content.InnerHtml; 
                    disc.Posts.Add(p);
                }
            }

            return disc;
        }

        public class DiscussionSet
        {
            public List<Discussion> Discussions { get; set; }
        }

        public class Discussion
        {
            [XmlAttribute]
            public string Title { get; set; }
            public List<Post> Posts { get; private set; } = new List<Post>();
        }

        public struct Post
        {
            [XmlAttribute]
            public string From { get; set; }
            [XmlAttribute]
            public string When { get; set; }
            public string Html { get; set; }
        }

        private static Task<T> StartSTATask<T>(Func<T> func)
        {
            var tcs = new TaskCompletionSource<T>();
            Thread thread = new Thread(() =>
            {
                try
                {
                    tcs.SetResult(func());
                }
                catch (Exception e)
                {
                    tcs.SetException(e);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return tcs.Task;
        }
    }
}
