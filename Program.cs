using System;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Globalization;



namespace fodt2tex
{
    class Border
    {
        public bool left = false;
        public bool right = false;
        public bool top = false;
        public bool bottom = false;
        public Border()
        {

        }
        public Border(Border ini)
        {
            this.left = ini.left;
            this.right = ini.right;
            this.top = ini.top;
            this.bottom = ini.bottom;
        }
    }
    class Program
    {

        static XNamespace fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        static XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        static XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        static XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

        static string GetText(XElement tcel)
        {
            if (tcel.IsEmpty)
                return "";
            string ttcel = "";
            var tts = tcel.Descendants(text + "span").Select(u => u.Value);
            foreach (string ts in tts)
            {
                if (!string.IsNullOrEmpty(ts))
                    ttcel += " " + ts;
            }
            /*
            if (string.IsNullOrEmpty(ttcel))
            {
                tts = tcel.Descendants(text + "p").Select(u => u.Value);
                foreach (string ts in tts)
                {
                    if (!string.IsNullOrEmpty(ts))
                        ttcel += ts;
                }
            }
            */
            ttcel = "\\scriptsize " + ttcel.Replace("_", "\\_").Replace("%", "\\%");
            return ttcel;
        }

        static void Main(string[] args)
        {
            if (args.Length == 0)
                return;
            string FullName = args[0];
            string pth = Path.GetDirectoryName(FullName);
            string fname = Path.GetFileNameWithoutExtension(FullName);
            string fhead = Path.Combine(pth, "head.t");
            string ftex = Path.Combine(pth, fname + ".tex");
            string texRes = File.ReadAllText(fhead);
            string mdoc = File.ReadAllText(FullName);


            XDocument xmdoc = XDocument.Parse(mdoc);

            //Отступы страницы
            var gml = xmdoc.Root.Descendants(style + "page-layout-properties").Select(u =>
            {
                if (u.Attribute(fo + "page-height") != null)
                {
                    string gm = $"\n\\usepackage[paperheight={u.Attribute(fo + "page-height").Value},paperwidth={u.Attribute(fo + "page-width").Value},left={u.Attribute(fo + "margin-left").Value},right={u.Attribute(fo + "margin-right").Value},top={u.Attribute(fo + "margin-top").Value},bottom={u.Attribute(fo + "margin-bottom").Value}]" + "{geometry}";
                    return gm;
                }
                else
                    return "";
            });
            string gm = "";
            foreach (string s in gml)
            {
                if (!string.IsNullOrEmpty(s))
                {
                    gm = s;
                    break;
                }
            }

            texRes += gm;
            texRes += "\n\\begin{document}";

            //Ширина столбцов
            Dictionary<string, string> columns = new Dictionary<string, string>();
            var cl = xmdoc.Root.Descendants(style + "style").Where(u => u.Attribute(style + "family").Value == "table-column").Select(
                u =>
                {
                    XElement p = u.Element(style + "table-column-properties");
                    KeyValuePair<string, string> r = KeyValuePair.Create(u.Attribute(style + "name").Value, p.Attribute(style + "column-width").Value);
                    return r;
                }
            );
            foreach (var e in cl)
            {
                columns.Add(e.Key, e.Value);
            }
            //Границы ячеек
            Dictionary<string, Border> cells = new Dictionary<string, Border>();
            var cel = xmdoc.Root.Descendants(style + "style").Where(u =>
            (u.Attribute(style + "family").Value == "table-cell")
            ).Select(
                u =>
                {
                    string tag = u.Attribute(style + "name").Value;
                    XElement p = u.Element(style + "table-cell-properties");
                    XAttribute d = new XAttribute("border", "none");
                    Border b = new Border();
                    if (p.Attribute(fo + "border") == null)
                    {
                        b.left = (p.Attribute(fo + "border-left") ?? d).Value != "none";
                        b.right = (p.Attribute(fo + "border-right") ?? d).Value != "none";
                        b.bottom = (p.Attribute(fo + "border-bottom") ?? d).Value != "none";
                        b.top = (p.Attribute(fo + "border-top") ?? d).Value != "none";
                    }
                    else
                    {
                        b.left = (p.Attribute(fo + "border").Value != "none");
                        b.right = (p.Attribute(fo + "border").Value != "none");
                        b.bottom = (p.Attribute(fo + "border").Value != "none");
                        b.top = (p.Attribute(fo + "border").Value != "none");
                    }
                    KeyValuePair<string, Border> r = KeyValuePair.Create(tag, b);
                    return r;
                }
            );
            foreach (var e in cel)
            {
                cells.Add(e.Key, e.Value);
            }

            var tables = xmdoc.Root.Descendants(table + "table").Select(u => u);
            XAttribute repe = new XAttribute("number-columns-repeated", "1");
            foreach (XElement tab in tables)
            {
                //нашли таблицу
                texRes += "\n\\begin{table}[H]\n\\begin{adjustbox}{max width=\\textwidth}\n\\begin{tabular}{";
                List<string> licol = new List<string>();
                List<decimal> wdcol = new List<decimal>();
                //Список колонок
                foreach (XElement tcol in tab.Elements(table + "table-column"))
                {
                    string tag = tcol.Attribute(table + "style-name").Value;
                    string vtag = columns[tag];
                    string wtag = "p{" + vtag + "}";

                    vtag = vtag.Replace("cm", "").Replace(".", ",");
                    decimal wd = decimal.Parse(vtag);
                    string sn = (tcol.Attribute(table + "number-columns-repeated") ?? repe).Value;
                    int n = int.Parse(sn);
                    for (int i = 0; i < n; i++)
                    {
                        licol.Add(wtag);
                        wdcol.Add(wd);
                    }
                }
                string tabul = string.Join("", licol.ToArray());
                texRes += tabul + "}";
                int ncol = wdcol.Count;
                decimal[] wds = wdcol.ToArray();
                for (int i = 1; i < ncol; i++)
                    wds[i] = wds[i] + wds[i - 1];

                //список строк
                string up_line = new string('~', ncol);
                string dn_line = new string('~', ncol);
                string sp_line = new string('~', ncol);
                Dictionary<KeyValuePair<int, int>, Border> sprow = new Dictionary<KeyValuePair<int, int>, Border>();
                int irow = 0;
                foreach (XElement trow in tab.Elements(table + "table-row"))
                {
                    string texrow = "";
                    int cpos = 0;
                    int ipos = 0;
                    char[] up_arr = up_line.ToCharArray();
                    char[] dn_arr = dn_line.ToCharArray();
                    //Формируем строку
                    //foreach (XElement tcel in trow.Elements(table + "table-cell"))

                    foreach (XElement tcel in trow.Elements())
                    {
                        if (ipos < cpos)
                        {
                            ipos++;
                            continue;
                        }
                        if (cpos >= ncol)
                            break;
                        string texcel = "\n\\multicolumn{";
                        Border b = new Border();
                        if (tcel.Attribute(table + "style-name") != null)
                        {
                            string tag = tcel.Attribute(table + "style-name").Value;
                            b = cells[tag];
                        }
                        else
                        {
                            //Это пустая ячейка, границы ищем по другому
                            KeyValuePair<int, int> ftag = KeyValuePair.Create(irow, ipos);
                            if (sprow.ContainsKey(ftag))
                                b = sprow[ftag];
                        }
                        string span = (tcel.Attribute(table + "number-columns-spanned") ?? repe).Value;

                        texcel += span + "}";

                        int nlen = int.Parse(span);
                        decimal lows = 0;
                        if (cpos > 0)
                            lows = wds[cpos - 1];
                        decimal ups = wds[cpos + nlen - 1];

                        string pw = "p{" + (ups - lows).ToString().Replace(",", ".") + "cm}";
                        if (b.left)
                            pw = "|" + pw;
                        if (b.right)
                            pw = pw + "|";

                        texcel += "{" + pw + "}";
                        //Здесь запишем текст
                        string ttcel = GetText(tcel);
                        texcel += "{" + ttcel + "}";

                        if (string.IsNullOrEmpty(texrow))
                            texrow = texcel;
                        else
                            texrow += " & " + texcel;
                        //Линии

                        if (b.top)
                        {
                            for (int i = cpos; i < cpos + nlen; i++)
                            {
                                up_arr[i] = '-';
                            }
                        }

                        if (b.bottom)
                        {
                            for (int i = cpos; i < cpos + nlen; i++)
                            {
                                dn_arr[i] = '-';
                            }
                        }

                        //rowspan
                        if (tcel.Attribute(table + "number-rows-spanned") != null)
                        {
                            for (int i = cpos; i < cpos + nlen; i++)
                            {
                                dn_arr[i] = '~';
                            }
                            int rwspan = int.Parse(tcel.Attribute(table + "number-rows-spanned").Value);
                            Border bl = new Border();
                            Border br = new Border();
                            Border bm = new Border();
                            bm.bottom = b.bottom;
                            bl.left = b.left;
                            br.right = b.right;
                            bl.bottom = false;
                            bl.top = false;

                            if (nlen == 1)
                            {
                                bl.right = b.right;
                                for (int i = irow + 1; i < irow + rwspan; i++)
                                {
                                    if (i == irow + rwspan - 1)
                                        bl.bottom = b.bottom;
                                    KeyValuePair<int, int> tg = KeyValuePair.Create(i, ipos);
                                    sprow.Add(tg, new Border(bl));
                                }
                            }
                            else
                            {
                                for (int i = irow + 1; i < irow + rwspan; i++)
                                {


                                    if (i == irow + rwspan - 1)
                                    {
                                        bl.bottom = b.bottom;
                                        br.bottom = b.bottom;
                                    }
                                    KeyValuePair<int, int> tg1 = KeyValuePair.Create(i, ipos);
                                    sprow.Add(tg1, new Border(bl));
                                    KeyValuePair<int, int> tg2 = KeyValuePair.Create(i, ipos + nlen - 1);
                                    sprow.Add(tg2, new Border(br));
                                }
                                //Нижняя граница
                                for (int i = ipos + 1; i < ipos + nlen - 1; i++)
                                {
                                    KeyValuePair<int, int> tg = KeyValuePair.Create(irow + rwspan - 1, i);
                                    sprow.Add(tg, new Border(bm));
                                }
                            }
                        }


                        cpos += nlen;
                        ipos++;
                    }

                    //сформировали texrow
                    up_line = new string(up_arr);
                    dn_line = new string(dn_arr);
                    if (up_line != sp_line)
                        texRes += "\n\\hhline{" + up_line + "}";
                    texRes += texrow + " \\\\ ";
                    up_line = dn_line;
                    dn_line = sp_line;
                    irow++;
                }
                //Добавляем последнюю черточку
                if (up_line != sp_line)
                    texRes += "\n\\hhline{" + up_line + "}";



                texRes += "\n\\end{tabular}\n\\end{adjustbox}\n\\end{table}";
                //break;
            }


            texRes += "\n\\end{document}";
            File.WriteAllText(ftex, texRes);

            ProcessStartInfo pi = new ProcessStartInfo();
            pi.WorkingDirectory = pth;
            pi.FileName = "xelatex";
            pi.Arguments = $"-interaction nonstopmode {ftex}";
            Process.Start(pi).WaitForExit();

            Console.WriteLine("successfully");

        }
    }
}
