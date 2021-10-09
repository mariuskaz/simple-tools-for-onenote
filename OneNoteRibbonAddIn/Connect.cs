using System;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.OneNote;
using OneNoteRibbonAddIn.Properties;
using System.Windows.Forms;
using System.Linq;
using System.Xml.Linq;
using System.Globalization;
using System.Net;
using Todoist.Net;
using System.Text.RegularExpressions;
using System.Net.Http;
using System.Collections.Generic;
using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace OneNoteRibbonAddIn
{
    [GuidAttribute("797efb51-6568-40c2-9564-f60683251281"), ProgId("OneNoteRibbonAddIn.Connect")]
    public class Connect : IRibbonExtensibility, IDTExtensibility2
    {
        private object _applicationObject;

        public string GetCustomUI(string ribbonId)
        {
            return Resources.customUI;
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = application;
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }        
        
        // This gets the image for the addin
        public IStream OnGetImage(string imageName)
        {
            MemoryStream stream = new MemoryStream();
            if (imageName == "todoist.png")
            {
                Resources.todoist.Save(stream, ImageFormat.Png);
            }

            return new ReadOnlyIStreamWrapper(stream);
        }

        public void InsertMonth(IRibbonControl control)
        {
            MessageBox.Show(DateTime.Now.ToString("MM"));
        }

        DateTime ganttStart;
        XElement gantt;
        XDocument doc;
        XNamespace ns;
        String style;
        
        public void SimpleGantt(IRibbonControl control)
        {
         
            String xml;
            Microsoft.Office.Interop.OneNote.Application onenote = new Microsoft.Office.Interop.OneNote.Application();
            string thisNoteBook = onenote.Windows.CurrentWindow.CurrentNotebookId;
            string thisSection = onenote.Windows.CurrentWindow.CurrentSectionId;
            string thisPage = onenote.Windows.CurrentWindow.CurrentPageId;
            onenote.GetPageContent(thisPage, out xml);

            doc = XDocument.Parse(xml);
            ns = doc.Root.Name.Namespace;
            style = "font-family:Calibri;font-size:9.0pt;";

            var gantts = from oe in doc.Descendants(ns + "OE")
                             from item in oe.Elements(ns + "Meta")
                                where item.Attribute("name").Value == "SimpleGanttTable"
                                    select oe;
 
            if (gantts.Count() == 0)
            {
                var outline = new XElement(ns + "Outline",
                    new XElement(ns + "Position",
                        new XAttribute("x", "36.0"),
                        new XAttribute("y", "80.0")
                    ),
                    new XElement(ns + "Size",
                        new XAttribute("width", "600.0"),
                        new XAttribute("height", "100.0"),
                        new XAttribute("isSetByUser", "true")
                    ),
                    new XElement(ns + "OEChildren",
                        new XElement(ns + "OE",
                            new XElement(ns + "Meta",
                                new XAttribute("name", "SimpleGanttStart"),
                                new XAttribute("content", "")
                            ),
                            new XElement(ns + "T", new XCData("Startas: ")) //DateTime.Now.ToString("yyyy.MM.dd")
                        ),
                        new XElement(ns + "OE",
                        new XElement(ns + "Meta",
                                new XAttribute("name", "SimpleGanttFinish"),
                                new XAttribute("content", "")
                            ),
                            new XElement(ns + "T", new XCData("Terminas: "))
                        ),
                        new XElement(ns + "OE",
                            new XElement(ns + "T", new XCData(""))
                        ),
                        new XElement(ns + "OE",
                            new XElement(ns + "Meta",
                                new XAttribute("name", "SimpleGanttTable"),
                                new XAttribute("content", "")
                            ),
                            new XElement(ns + "Table",
                                new XAttribute("bordersVisible", "true"),
                                new XAttribute("hasHeaderRow", "true"),

                                new XElement(ns + "Columns",
                                    new XElement(ns + "Column",
                                        new XAttribute("index", "0"),
                                        new XAttribute("width", "140.0"),
                                        new XAttribute("isLocked", "true")
                                    ),
                                    new XElement(ns + "Column",
                                        new XAttribute("index", "1"),
                                        new XAttribute("width", "40.0"),
                                        new XAttribute("isLocked", "true")
                                    ),
                                    new XElement(ns + "Column",
                                        new XAttribute("index", "2"),
                                        new XAttribute("width", "40.0"),
                                        new XAttribute("isLocked", "true")
                                    ),
                                    new XElement(ns + "Column",
                                        new XAttribute("index", "3"),
                                        new XAttribute("width", "80.0"),
                                        new XAttribute("isLocked", "true")
                                    )
                                ),

                                new XElement(ns + "Row",
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XAttribute("style", style),
                                                new XAttribute("alignment", "center"),
                                                new XElement(ns + "T", new XCData("UŽDUOTIS"))
                                            )
                                        )
                                    ),
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XAttribute("style", style),
                                                new XAttribute("alignment", "center"),
                                                new XElement(ns + "T", new XCData("STARTAS"))
                                            )
                                        )
                                    ),
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XAttribute("style", style),
                                                new XAttribute("alignment", "center"),
                                                new XElement(ns + "T", new XCData("TRUKMĖ"))
                                            )
                                        )
                                    ),
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XAttribute("style", style),
                                                new XAttribute("alignment", "center"),
                                                new XElement(ns + "T", new XCData("KAS?"))
                                            )
                                        )
                                    )
                                ),

                                new XElement(ns + "Row",
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XElement(ns + "T", new XCData("Užduotis 1"))
                                            )
                                        )
                                    ),
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XElement(ns + "T", new XCData("1"))
                                            )
                                        )
                                    ),
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XElement(ns + "T", new XCData("1"))
                                            )
                                        )
                                    ),
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XElement(ns + "T", new XCData(""))
                                            )
                                        )
                                    )
                                ),


                                new XElement(ns + "Row",
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XElement(ns + "T", new XCData(""))
                                            )
                                        )
                                    ),
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XElement(ns + "T", new XCData(""))
                                            )
                                        )
                                    ),
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XElement(ns + "T", new XCData(""))
                                            )
                                        )
                                    ),
                                    new XElement(ns + "Cell",
                                        new XElement(ns + "OEChildren",
                                            new XElement(ns + "OE",
                                                new XElement(ns + "T", new XCData(""))
                                            )
                                        )
                                    )
                                )


                            )
                        )
                    )
                );

                var page = doc.Descendants(ns + "Page").First();
                page.Add(outline);

                gantts = from oe in doc.Descendants(ns + "OE")
                         from item in oe.Elements(ns + "Meta")
                         where item.Attribute("name").Value == "SimpleGanttTable"
                         select oe;

            }

            gantt = gantts.ElementAt(0);

            var dates = from oe in doc.Descendants(ns + "OEChildren")
                        from item in oe.Descendants(ns + "Meta")
                        where item.Attribute("name").Value == "SimpleGanttTable"
                        select oe; //.Descendants(ns + "T");

            if (dates.Count() > 0)
            {
                var startTag = RemoveHtmlTags(dates.Descendants(ns+"T").First().Value).Substring(8).Trim();
                DateTime.TryParse(startTag, out ganttStart);
            }
            
            // CALC COLUMNS //

            int taskColumn = 1;
            int startColumn = 2;
            int durationColumn = 3;
            string taskName = "";
            int start = 0;
            int duration = 0;
            int maxPeriod = 0;
            string[] weekdays = { "VII", "I", "II", "III", "IV", "V", "VI" };

            var items = gantt.Elements(ns + "Table").First().Descendants(ns + "Cell");
            int cols = gantt.Elements(ns + "Table").First().Descendants(ns + "Column").Count();
            int _cols = cols;
            int col = 0;

            foreach (var item in items)
            {
                col++;
                if (col == startColumn)
                {
                    if (validDate(item.Value))
                    {
                        var startDate = DateTime.Parse(item.Value);
                        if (ganttStart == new DateTime()) ganttStart = startDate;
                        var timespan = startDate - ganttStart;
                        start = Convert.ToInt32(timespan.TotalDays) + 1;
                    }
                    else
                    {
                        Int32.TryParse(item.Value, out start);
                    }
                }
                if (col == durationColumn & start > 0 & Int32.TryParse(item.Value, out duration))
                {
                    if (start + duration - 1 > maxPeriod) maxPeriod = start + duration - 1;
                }
                if (col == _cols) col = 0;
            }

            // ADD COLUMNS //

            addGanttColumns(maxPeriod + 4 - cols);
            cols = gantt.Elements(ns + "Table").First().Descendants(ns + "Column").Count();


            // SET HEADER ROW //

            var headers = gantt.Descendants(ns + "Row").First().Descendants(ns + "Cell");
            foreach (var header in headers)
            {
                col++;
                if (col > 4)
                {
                    var index = col - 4;
                    String txt = index.ToString();
                    if (ganttStart != new DateTime())
                    {
                        var date = ganttStart.AddDays(col - 5);
                        var weekday = (int)date.DayOfWeek;
                        txt = Right("0" + date.Month.ToString(), 2) + "." + Right("0" + date.Day.ToString(), 2);
                        txt = txt + System.Environment.NewLine + weekdays[weekday];
                    }
                    header.Descendants(ns + "T").First().Value = txt;
                }
            }
            col = 0;


            // ADD COLORS //

            var cells = gantt.Elements(ns + "Table").First().Descendants(ns + "Cell");
            foreach (var cell in cells)
            {
                col++;
                if (col == taskColumn) taskName = cell.Value;
                if (col == startColumn)
                {
                    if (validDate(cell.Value))
                    {
                        var startDate = DateTime.Parse(cell.Value);
                        if (ganttStart == new DateTime()) ganttStart = startDate;
                        var timespan = startDate - ganttStart;
                        start = Convert.ToInt32(timespan.TotalDays) + 1;
                    }
                    else
                    {
                        Int32.TryParse(cell.Value, out start);
                    }
                }
                if (col == durationColumn) Int32.TryParse(cell.Value, out duration);
                var color = cell.Attribute("shadingColor");
                if (color != null) cell.Attribute("shadingColor").Remove();
                var finish = start + duration - 1;
                var current = col - 4;

                var weekend = false;
                if (cell.Value.IndexOf("VI") > 0) weekend = true;

                if (col > 4 & current >= start & current <= finish)
                {
                    if (taskName == taskName.ToUpper())
                    {
                        cell.Add(new XAttribute("shadingColor", "#5F497A"));
                    }
                    else
                    {
                        cell.Add(new XAttribute("shadingColor", "#CCC1D9"));
                    }
                }
                else if (col > 4)
                {
                    if (weekend == true)
                    {
                        cell.Add(new XAttribute("shadingColor", "#D6DCE4"));
                    }
                    else
                    {
                        if (col % 2 > 0) cell.Add(new XAttribute("shadingColor", "#FAFAFA"));
                    }
                }
                if (col == cols) col = 0;
            }

            //doc.Save("D:/doc.xml");
            onenote.UpdatePageContent(doc.ToString());

        }

        public void addGanttColumns(int cols)
        {
            for (var c = 0; c < cols; c++)
            {
                int col = gantt.Elements(ns + "Table").First().Descendants(ns + "Column").Count() - 3;
                int index = col + 3;
                var columns = gantt.Elements(ns + "Table").First().Descendants(ns + "Columns").First();
                var column = new XElement(ns + "Column",
                    new XAttribute("index", index.ToString()),
                    new XAttribute("width", "10.0"),
                    new XAttribute("isLocked", "false")
                );
                columns.Add(column);

                var rows = gantt.Elements(ns + "Table").First().Descendants(ns + "Row");
                foreach (var row in rows)
                {
                    var cell = new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                            new XElement(ns + "OE",
                                new XAttribute("style", style),
                                new XAttribute("alignment", "center"),
                                new XElement(ns + "T", new XCData(""))
                            )
                        )
                    );
                    row.Add(cell);
                }

            }
        }

        String todoist_user = "";
        String todoist_psw = "";
        
        public async void ExportTasks(IRibbonControl control)
        {

            Window context = control.Context as Window;
            CWin32WindowWrapper owner = new CWin32WindowWrapper((IntPtr)context.WindowHandle);

            Microsoft.Office.Interop.OneNote.Application onenote = new Microsoft.Office.Interop.OneNote.Application();
            string thisNoteBook = onenote.Windows.CurrentWindow.CurrentNotebookId;
            string thisSection = onenote.Windows.CurrentWindow.CurrentSectionId;
            string thisPage = onenote.Windows.CurrentWindow.CurrentPageId;

            String link;
            onenote.GetHyperlinkToObject(thisPage, System.String.Empty, out link);

            String xmlNotebooks;
            String xmlPage;

            onenote.GetHierarchy(null,
               Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, out xmlNotebooks);

            var notebooks = XDocument.Parse(xmlNotebooks);
            var currentbook = from item in notebooks.Descendants(notebooks.Root.Name.Namespace + "Notebook")
                              where item.Attribute("ID").Value == thisNoteBook
                              select item;

            var notebook = currentbook.First().Attribute("name").Value;
            onenote.GetPageContent(thisPage, out xmlPage);
            doc = XDocument.Parse(xmlPage);
            ns = doc.Root.Name.Namespace;

            var title = RemoveHtmlTags(doc.Descendants(ns + "Title").First().Value);
            var todolist = new List<Todo>();

            var gantt = (from oe in doc.Descendants(ns + "OE")
                     from item in oe.Elements(ns + "Meta")
                     where item.Attribute("name").Value == "SimpleGanttTable"
                     select oe).FirstOrDefault();

            if (gantt != null)
            {
                var rows = gantt.Elements(ns + "Table").Descendants(ns + "Row");
                foreach (var row in rows)
                {
                    var cells = row.Descendants(ns + "Cell");
                    var tags = from cell in cells
                               from tag in cell.Descendants(ns + "Tag")
                               where tag.Attribute("completed").Value == "false"
                               select cell;
                    if (tags.Count() > 0)
                        todolist.Add(new Todo { content = cells.ElementAt(0).Value, assignedTo = cells.ElementAt(3).Value, due = cells.ElementAt(1).Value });
                }
            }

            if (todolist.Count() == 0)
            {
                var tasks = from oe in doc.Descendants(ns + "OE")
                            from item in oe.Elements(ns + "Tag")
                            where item.Attribute("completed").Value == "false"
                            select oe;

                if (tasks.Count() == 0)
                {
                    MessageBox.Show(owner, "No tasks found on this page!", title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                foreach (var task in tasks)
                    todolist.Add(new Todo { content = task.Value });

            }


            // LOGIN IF NOT //
            if (todoist_user.Length == 0)
            {
                LoginForm login = new LoginForm();
                login.ShowDialog(owner);
                if (login.DialogResult == DialogResult.OK)
                {
                    if (login.email.Contains("@") == false) login.email += "@ardi.lt";
                    todoist_user = login.email;
                    todoist_psw = login.password;
                    login.Dispose();
                    login = null;
                }
            }


            if (todoist_user.Length > 0)
            {

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                ITodoistTokenlessClient tokenlessClient = new TodoistTokenlessClient();

                try
                {

                    ITodoistClient client = await tokenlessClient.LoginAsync(todoist_user, todoist_psw);
                    var projects = await client.Projects.GetAsync();
                    var status = "Tasks found: " + todolist.Count().ToString();

                    TasksForm confirm = new TasksForm(title, status, projects);
                    confirm.ShowDialog(owner);

                    if (confirm.DialogResult == DialogResult.OK)
                    {

                        var transaction = client.CreateTransaction();
                        var user = await client.Users.GetCurrentAsync();
                        var token = user.Token.ToString();

                        // Get all collaborators //
                        var userDetails = new System.Collections.Hashtable();
                        using (var httpClient = new HttpClient())
                        {
                            using (var request = new HttpRequestMessage(new HttpMethod("POST"), "https://api.todoist.com/sync/v8/sync"))
                            {
                                request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + token);
                                var contentList = new List<string>();
                                contentList.Add("sync_token=*");
                                contentList.Add("resource_types=[\"collaborators\"]");
                                request.Content = new StringContent(string.Join("&", contentList));
                                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/x-www-form-urlencoded");
                                var response = await httpClient.SendAsync(request);
                                var json = response.Content.ReadAsStringAsync();
                                var result = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(json.Result);
                                var persons = result.collaborators;
                                foreach (var person in persons)
                                {
                                    string email = person.email;
                                    string mail = email.ToLower();
                                    long userId = person.id;
                                    userDetails.Add(mail, userId);
                                }
                            }
                        }

                        var projectId = new Todoist.Net.Models.ComplexId();
                        if (confirm.id > 0)
                        {
                            foreach (var item in projects)
                            {
                                if (item.Name == confirm.project) projectId = item.Id;
                            }
                        }

                        else
                        {
                            projectId = await transaction.Project.AddAsync(new Todoist.Net.Models.Project(confirm.project));
                            //await transaction.Sharing.ShareProjectAsync(projectId, "pasitarimai@ardi.lt");
                        }
               

                        foreach (var task in todolist)
                        {
                            string[] text = task.content.ToString().Split('#');
                            var content = RemoveHtmlTags(text[0]);
                            if (confirm.links) content = "[" + content + "](" + link + ")";

                            Todoist.Net.Models.Item todo = 
                                new Todoist.Net.Models.Item(content, projectId);

                            if (validDate(task.due))
                                todo.DueDate = new Todoist.Net.Models.DueDate(task.due);

                            if (task.assignedTo.Length > 0)
                            {
                                string mail = task.assignedTo.IndexOf("@") > 0 ? task.assignedTo : task.assignedTo + "@ardi.lt";
                                string key = mail.ToLower();
                                if (userDetails.ContainsKey(key))
                                {
                                    long userId= (long)userDetails[key];
                                    todo.ResponsibleUid = userId;
                                    await transaction.Sharing.ShareProjectAsync(projectId, mail);
                                }
                            }

                            var taskId = await transaction.Items.AddAsync(todo);
                            if (text.Length > 1) await transaction.Notes.AddToItemAsync(new Todoist.Net.Models.Note(text[1]), taskId);     
                        }

                        await transaction.CommitAsync();
                        string url = "https://todoist.com";
                        projects = await client.Projects.GetAsync();
                        foreach (var project in projects)
                        {
                            if (project.Name == confirm.project 
                                && !project.IsArchived ) url = url + "/app/project/" + project.Id.ToString();
                        }
                        System.Diagnostics.Process.Start(url);
                    }

                    confirm.Dispose();
                    confirm = null;

                }
                catch
                {
                    MessageBox.Show(owner, "Bad user mail or password...", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            
        }

        public class Todo {
            public string content { get; set; }
            public string assignedTo { get; set; }
            public string due { get; set; }
        }

        public class User
        {
            public long id { get; set; }
            public string full_name { get; set; }
            public string email { get; set; }
        }

        private string Right(string str, int x)
        {
            return str.Substring(str.Length - x);
        }

        private string RemoveHtmlTags(string html)
        {
            return Regex.Replace(html, @"<(.|\n)*?>", "");
        }

        private Boolean validDate(string date)
        {
            Regex regex = new Regex(@"\d\d\d\d-\d\d-\d\d");
            return regex.IsMatch(date);
        }

    }
}
