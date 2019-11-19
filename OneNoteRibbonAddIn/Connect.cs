﻿using System;
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

namespace OneNoteRibbonAddIn
{
    [GuidAttribute("797efb51-6568-40c2-9564-f60683251281"), ProgId("OneNoteRibbonAddIn.Connect")]
    public class Connect : IRibbonExtensibility, IDTExtensibility2
    {
        private object _applicationObject;
        private String API_token = "c011a6914a334f7fba3f4b4b0d088bb5382eb790";

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

        XDocument doc;
        XNamespace ns;
        XElement gantt;
        DateTime ganttStart;
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
                            new XElement(ns + "T", new XCData("Startas: " + DateTime.Now.ToString("yyyy.MM.dd")))
                        ),
                        new XElement(ns + "OE",
                        new XElement(ns + "Meta",
                                new XAttribute("name", "SimpleGanttFinish"),
                                new XAttribute("content", "")
                            ),
                            new XElement(ns + "T", new XCData("Terminas: " + DateTime.Now.ToString("yyyy.MM.dd")))
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
            ganttStart = new DateTime();

            var dates = from oe in doc.Descendants(ns + "OE")
                     from item in oe.Elements(ns + "Meta")
                     where item.Attribute("name").Value == "SimpleGanttStart"
                     select oe;

            if (dates.Count() > 0)
                DateTime.TryParse(dates.ElementAt(0).Value.Substring(8), out ganttStart);

            // CALC COLUMNS //

            int startColumn = 2;
            int durationColumn = 3;
            int start = 0;
            int duration = 0;
            int maxPeriod = 0;

            var items = gantt.Elements(ns + "Table").First().Descendants(ns + "Cell");
            int cols = gantt.Elements(ns + "Table").First().Descendants(ns + "Column").Count();
            int _cols = cols;
            int col = 0;

            foreach (var item in items)
            {
                col++;
                if (col == startColumn) Int32.TryParse(item.Value, out start);
                if (col == durationColumn & start > 0 & Int32.TryParse(item.Value, out duration))
                {
                    if (start + duration - 1 > maxPeriod) maxPeriod = start + duration - 1;
                }
                if (col == _cols) col = 0;
            }

            // ADD COLUMNS//

            addGanttColumns(maxPeriod + 4 - cols);
            cols = gantt.Elements(ns + "Table").First().Descendants(ns + "Column").Count();

            // ADD COLORS //

            var cells = gantt.Elements(ns + "Table").First().Descendants(ns + "Cell");
            foreach (var cell in cells)
            {
                col++;
                if (col == startColumn) Int32.TryParse(cell.Value, out start);
                if (col == durationColumn) Int32.TryParse(cell.Value, out duration);
                var color = cell.Attribute("shadingColor");
                if (color != null) cell.Attribute("shadingColor").Remove();
                var finish = start + duration - 1;
                var current = col - 4;
                if (col > 4 & current >= start & current <= finish)
                {
                    cell.Add(new XAttribute("shadingColor", "#CCC1D9"));
                }
                else if (col > 4)
                {
                    if (col % 2 > 0) cell.Add(new XAttribute("shadingColor", "#FAFAFA"));
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
                    new XAttribute("width", "20.0"),
                    new XAttribute("isLocked", "true")
                );
                columns.Add(column);

                var rows = gantt.Elements(ns + "Table").First().Descendants(ns + "Row");

                String txt = col.ToString();
                if (ganttStart != new DateTime())
                {
                    var date = ganttStart.AddDays(col - 1);
                    txt = date.Month.ToString() + "." + date.Day.ToString();
                }

                foreach (var row in rows)
                {
                    var cell = new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                            new XElement(ns + "OE",
                                new XAttribute("style", style),
                                new XAttribute("alignment", "center"),
                                new XElement(ns + "T", new XCData(txt))
                            )
                        )
                    );
                    row.Add(cell);
                    txt = " ";
                }

            }
        }



        public async void ExportTasks(IRibbonControl control)
        {
            String xml;
            Microsoft.Office.Interop.OneNote.Application onenote = new Microsoft.Office.Interop.OneNote.Application();
            string thisNoteBook = onenote.Windows.CurrentWindow.CurrentNotebookId;
            string thisSection = onenote.Windows.CurrentWindow.CurrentSectionId;
            string thisPage = onenote.Windows.CurrentWindow.CurrentPageId;
            onenote.GetPageContent(thisPage, out xml);

            doc = XDocument.Parse(xml);
            ns = doc.Root.Name.Namespace;

            var title = doc.Descendants(ns + "Title").First().Value;

            var tasks = from oe in doc.Descendants(ns + "OE")
                       from item in oe.Elements(ns + "Tag")
                       where item.Attribute("index").Value == "0"
                       where item.Attribute("completed").Value == "false"
                       select oe;

            String info = title + "\n\n";
            foreach (var task in tasks)
            {
                info = info + "o  " + task.Value + "\n";

            }

            string caption = "Add tasks to Todoist:";
            MessageBoxButtons buttons = MessageBoxButtons.OKCancel;
            DialogResult result;

            result = MessageBox.Show(info, caption, buttons);

            if (result == DialogResult.OK)
            {

                //https://github.com/olsh/todoist-net //

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                ITodoistClient client = new TodoistClient(API_token);

                var transaction = client.CreateTransaction();
                var projectId = await transaction.Project.AddAsync(new Todoist.Net.Models.Project(title));
                foreach (var task in tasks)
                {
                    var taskId = await transaction.Items.AddAsync(new Todoist.Net.Models.Item(task.Value, projectId));
                    //await transaction.Notes.AddToItemAsync(new Todoist.Net.Models.Note("Task description"), taskId);
                }

                await transaction.CommitAsync();




                System.Diagnostics.Process.Start("https://todoist.com");

            }



        }





    }
}
