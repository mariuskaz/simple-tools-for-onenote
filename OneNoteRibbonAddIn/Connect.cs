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
            if (imageName == "showform.png")
            {
                Resources.showform.Save(stream, ImageFormat.Png);
            }

            return new ReadOnlyIStreamWrapper(stream);
        }

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
            doc.Save("D:/one.xml");

            // OUTLINE //
            if (doc.Descendants(ns + "Table").Count() == 0)
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
                            new XElement(ns + "T", new XCData("Startas: " + DateTime.Now.ToString("yyyy.MM.dd")))
                        ),
                        new XElement(ns + "OE",
                            new XElement(ns + "T", new XCData("Pabaiga: " + DateTime.Now.ToString("yyyy.MM.dd")))
                        ),
                        new XElement(ns + "OE",
                            new XElement(ns + "T", new XCData(""))
                        ),
                        new XElement(ns + "OE",
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
                                                new XElement(ns + "T", new XCData("Uzregistruoti MS"))
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

            }
            
            int cols = doc.Descendants(ns + "Table").First().Descendants(ns+"Column").Count();
            int _cols = cols;

            int startColumn = 2;
            int durationColumn = 3;
            int col = 0;
            int start = 0;
            int duration = 0;
            int maxPeriod = 0;

            var items = doc.Descendants(ns + "Table").First().Descendants(ns + "Cell");
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
            
            addColumns(maxPeriod + 4 - cols);
            cols = doc.Descendants(ns + "Table").First().Descendants(ns + "Column").Count();

            var cells = doc.Descendants(ns + "Table").First().Descendants(ns + "Cell");
            foreach (var cell in cells)
            {

            }

            doc.Save("D:/doc.xml");
            onenote.UpdatePageContent(doc.ToString());

        }

        public void addColumns(int cols)
        {
            for (var c = 0; c < cols; c++)
            {
                int col = doc.Descendants(ns + "Table").First().Descendants(ns + "Column").Count() - 3;
                var columns = doc.Descendants(ns + "Table").First().Descendants(ns + "Columns").First();
                var column = new XElement(ns + "Column",
                    new XAttribute("index", "4"),
                    new XAttribute("width", "20.0"),
                    new XAttribute("isLocked", "true")
                );
                columns.Add(column);

                var rows = doc.Descendants(ns + "Table").First().Descendants(ns + "Row");
                var r = 0;
                var txt = col.ToString();
                
                foreach (var row in rows)
                {
                    r++;
                    if (r > 1) txt = "";
                    var cell = new XElement(ns + "Cell",
                        new XElement(ns + "OEChildren",
                            new XElement(ns + "OE",
                                new XAttribute("style", style),
                                new XAttribute("alignment", "center"),
                                new XElement(ns + "T", new XCData(txt))
                            )
                        )
                    );
                    if (col % 2 > 0) cell.Add(new XAttribute("shadingColor", "#FAFAFA"));
                    row.Add(cell);
                }

            }
        }





    }
}