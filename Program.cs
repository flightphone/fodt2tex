using System;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml;

using System.Globalization;
using System.Drawing;
using System.Drawing.Imaging;



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
        static XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

        static XNamespace svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
        static XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

        static Dictionary<string, string> columns = new Dictionary<string, string>();
        static Dictionary<string, Border> cells = new Dictionary<string, Border>();

        static Dictionary<string, KeyValuePair<string, string>> styles = new Dictionary<string, KeyValuePair<string, string>>();
        static string[] fontsize;
        static KeyValuePair<string, string> getFont(XElement t)
        {
            string start = "";
            string end = "";
            if (t.Attribute(fo + "font-size") != null)
            {
                string sz = t.Attribute(fo + "font-size").Value;
                try
                {
                    decimal dic = decimal.Parse(t.Attribute(fo + "font-size").Value.Replace("pt", "").Replace(".", ","));
                    int ic = (int)dic;
                    string vtex = "HUGE";
                    if (ic <= 30)
                        vtex = fontsize[ic];

                    start += vtex;
                    end += "}";
                }
                catch
                {
                    sz = "";
                }

            }
            else
            if (t.Attribute(style + "font-size-complex") != null)
            {
                string sz = t.Attribute(style + "font-size-complex").Value;
                try
                {
                    decimal dic = decimal.Parse(t.Attribute(style + "font-size-complex").Value.Replace("pt", "").Replace(".", ","));
                    int ic = (int)dic;
                    string vtex = "HUGE";
                    if (ic <= 30)
                        vtex = fontsize[ic];
                    start += vtex;
                    end += "}";
                }
                catch
                {
                    sz = "";
                }
            }
            //bold italic
            if (t.Attribute(fo + "font-weight") != null)
            {
                if (t.Attribute(fo + "font-weight").Value == "bold")
                {
                    start += "\\textbf{";
                    end += "}";
                }
            }

            if (t.Attribute(fo + "font-style") != null)
            {
                if (t.Attribute(fo + "font-style").Value == "italic")
                {
                    start += "\\textit{";
                    end += "}";
                }
            }
            return KeyValuePair.Create(start, end);

        }
        static void initStyles(XDocument xmdoc)
        {
            //инициализация справочников
            var ttcl = xmdoc.Descendants(style + "style").Where(u => u.Attribute(style + "family").Value == "text").Select(
                u =>
                {

                    string tag = u.Attribute(style + "name").Value;
                    string start = "";
                    String end = "";

                    XElement t = u.Element(style + "text-properties");
                    if (t != null)
                    {
                        KeyValuePair<string, string> fn = getFont(t);
                        start += fn.Key;
                        end += fn.Value;
                    }

                    KeyValuePair<string, KeyValuePair<string, string>> r = KeyValuePair.Create(tag, KeyValuePair.Create(start, end));
                    return r;
                }
            );

            foreach (var e in ttcl)
            {
                if (!styles.ContainsKey(e.Key))
                    styles.Add(e.Key, e.Value);

            }

            var parcl = xmdoc.Descendants(style + "style").Where(u => u.Attribute(style + "family").Value == "paragraph").Select(
                u =>
                {

                    string tag = u.Attribute(style + "name").Value;
                    string start = "";
                    String end = "";
                    XElement p = u.Element(style + "paragraph-properties");
                    if (p != null)
                        if (p.Attribute(fo + "text-align") != null)
                        {
                            string v = p.Attribute(fo + "text-align").Value;
                            string vtex = "{";
                            if (v == "start")
                                vtex = "\\raggedright{";

                            if (v == "end")
                                vtex = "\\raggedleft{";

                            if (v == "center")
                                vtex = "\\centering{";

                            start += vtex;
                            end += "}";
                        }


                    XElement t = u.Element(style + "text-properties");
                    if (t != null)
                        if (t != null)
                        {
                            KeyValuePair<string, string> fn = getFont(t);
                            start += fn.Key;
                            end += fn.Value;
                        }



                    KeyValuePair<string, KeyValuePair<string, string>> r = KeyValuePair.Create(tag, KeyValuePair.Create(start, end));
                    return r;
                }
            );

            foreach (var e in parcl)
            {
                if (!styles.ContainsKey(e.Key))
                    styles.Add(e.Key, e.Value);
            }

            //Ширина столбцов
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
                if (!columns.ContainsKey(e.Key))
                    columns.Add(e.Key, e.Value);
            }
            //Границы ячеек
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
                if (!cells.ContainsKey(e.Key))
                    cells.Add(e.Key, e.Value);
            }
            //инициализация справочников

        }
        static string esctex(string s)
        {
            string md = "DE5BA24A-BCFC-4D5D-B92E-E283F9D39ABE";
            return s
            .Replace("\\", $"{md}\\backslash{md}")
            .Replace("~", $"{md}\\sim{md}")
            .Replace("_", "\\_")
            .Replace("%", "\\%")
            .Replace("&", "\\&")
            .Replace("$", "\\$")
            .Replace("{", "\\{")
            .Replace("}", "\\}")
            .Replace("#", "\\#")
            .Replace("^", md + "\\hat{}" + md)
            .Replace(md, "$");
        }

        static int nimage = 0;


        static string Gettt(XElement tcel)
        //только текст
        {
            if (tcel.Name == draw + "frame")
            {
                return "";
            }

            if (tcel.Name == text + "line-break")
            {
                //новая строка
                return "\t";
            }

            string start = "";  //Подгружаем из таблицы стилей
            string end = "";
            if (tcel.Name == text + "p")
            {
                end = "\t";
            }

            if (!tcel.HasElements)
            {
                if (string.IsNullOrEmpty(tcel.Value))
                    tcel.Value = " ";
                return start + tcel.Value + end;
            }
            else
            {
                string tts = start;
                foreach (XNode el in tcel.Nodes())
                {
                    if (el is XElement)
                        tts += Gettt((XElement)el);
                    else
                    if (el.NodeType == XmlNodeType.Text)
                    {
                        tts += el.ToString();
                    }

                }
                tts += end;
                return tts;
            }
        }

        static string GetText(XElement tcel)
        {
            if (tcel.Name == draw + "frame")
            {
                //Это картинка
                if (tcel.Attribute(svg + "width") == null)
                    return "";
                try
                {
                    string w = tcel.Attribute(svg + "width").Value;
                    string h = tcel.Attribute(svg + "height").Value;
                    string imgname = $"{fname}_img{nimage}.png";
                    string fullname = Path.Combine(pth, imgname);
                    string b64 = tcel.Element(draw + "image").Element(office + "binary-data").Value;
                    byte[] imageBytes = Convert.FromBase64String(b64);
                    using (var ms = new MemoryStream(imageBytes, 0, imageBytes.Length))
                    {
                        Image image = Image.FromStream(ms, true);
                        image.Save(fullname, ImageFormat.Png);
                        nimage++;
                    }
                    string res = $@"\includegraphics[width={w},height={h}]" + "{" + imgname + "}";
                    return res;
                }
                catch
                {
                    return "";
                }
            }

            if (tcel.Name == text + "line-break")
            {
                //новая строка
                return "\\par ";
            }

            string start = "";  //Подгружаем из таблицы стилей
            string end = "";
            if (tcel.Name == text + "p")
            {
                start = "\\par ";
            }
            if (tcel.Attribute(text + "style-name") != null)
            {
                string tag = tcel.Attribute(text + "style-name").Value;
                if (styles.ContainsKey(tag))
                {
                    KeyValuePair<string, string> se = styles[tag];
                    start = start += se.Key;
                    end = se.Value;
                }
            }

            if (!tcel.HasElements)
            {
                if (string.IsNullOrEmpty(tcel.Value))
                    tcel.Value = " ";
                return start + esctex(tcel.Value) + end;
            }
            else
            {
                string tts = start;
                foreach (XNode el in tcel.Nodes())
                {
                    if (el is XElement)
                        tts += GetText((XElement)el);
                    else
                    if (el.NodeType == XmlNodeType.Text)
                    {
                        tts += esctex(el.ToString());
                    }


                }
                tts += end;
                return tts;
            }
        }

        static string pth;
        static string fname;
        static void Main(string[] args)
        {
            if (args.Length == 0)
                return;
            string FullName = args[0];
            pth = Path.GetDirectoryName(FullName);
            fname = Path.GetFileNameWithoutExtension(FullName);
            string fhead = Path.Combine(pth, "head.t");
            string ffont = Path.Combine(pth, "font_size.t");
            string ftex = Path.Combine(pth, fname + ".tex");
            string ftxt = Path.Combine(pth, fname + ".txt"); //Паралелльно запишем тольео текст
            string texRes = File.ReadAllText(fhead);
            string txtRes = "";
            string mdoc = File.ReadAllText(FullName);

            fontsize = File.ReadAllLines(ffont); //Размеры шрифтов
            XDocument xmdoc = XDocument.Parse(mdoc);
            initStyles(xmdoc);

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


            var tables = xmdoc.Root.Descendants(table + "table").Select(u => u);
            XAttribute repe = new XAttribute("number-columns-repeated", "1");
            foreach (XElement tab in tables)
            {
                //нашли таблицу
                txtRes += "\n\n\n";

                texRes += "\n\\begin{table}[H]\n\\begin{adjustbox}{max width=\\textwidth}\n\\begin{tabular}{";
                List<string> licol = new List<string>();
                List<decimal> wdcol = new List<decimal>();
                //Список колонок
                decimal scale = 1.0M;
                if (tab.Attribute("scale") != null)
                {
                    scale = Decimal.Parse(tab.Attribute("scale").Value.Replace(".", ","));
                }


                foreach (XElement tcol in tab.Elements(table + "table-column"))
                {
                    string tag = tcol.Attribute(table + "style-name").Value;
                    string vtag = columns[tag];
                    vtag = vtag.Replace("cm", "").Replace(".", ",");
                    decimal wd = decimal.Parse(vtag) / scale;
                    string wtag = "p{" + wd.ToString().Replace(",", ".") + "}";
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
                    string ttrow = ""; //только текст
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

                        string pw = "{" + (ups - lows).ToString().Replace(",", ".") + "cm}";
                        string multirow = "";
                        string endcell = "";

                        if (tcel.Attribute(table + "number-rows-spanned") != null)
                        {
                            multirow = "{\\multirow{" + tcel.Attribute(table + "number-rows-spanned").Value
                            + "}{*}{\\parbox" + pw;
                            endcell = "}}";
                        }


                        pw = "p" + pw;
                        if (b.left)
                            pw = "|" + pw;
                        if (b.right)
                            pw = pw + "|";

                        texcel += "{" + pw + "}" + multirow;


                        //Здесь запишем текст
                        string ttcel = GetText(tcel);
                        texcel += "{" + ttcel + "}" + endcell;



                        if (string.IsNullOrEmpty(texrow))
                            texrow = texcel;
                        else
                            texrow += " & " + texcel;
                        //Линии

                        string tt = Gettt(tcel);
                        if (string.IsNullOrEmpty(ttrow))
                        {
                            if (tt[tt.Length - 1] != '\t')
                                tt = tt + "\t";
                            ttrow = tt;
                        }
                        else
                        {
                            if (tt[tt.Length - 1] != '\t')
                                tt = tt + "\t";
                            ttrow += tt;
                        }

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

                    txtRes += ttrow + "\n";
                }
                //Добавляем последнюю черточку
                if (up_line != sp_line)
                    texRes += "\n\\hhline{" + up_line + "}";



                texRes += "\n\\end{tabular}\n\\end{adjustbox}\n\\end{table}";
                //break;
            }


            texRes += "\n\\end{document}";
            File.WriteAllText(ftex, texRes);
            File.WriteAllText(ftxt, txtRes);

            ProcessStartInfo pi = new ProcessStartInfo();
            pi.WorkingDirectory = pth;
            pi.FileName = "xelatex";
            pi.Arguments = $"-interaction nonstopmode {ftex}";
            Process.Start(pi).WaitForExit();

            Console.WriteLine("successfully");

        }
    }
}
