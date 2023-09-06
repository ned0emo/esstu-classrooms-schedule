using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SheduleSI
{
    class Repository
    {
        public async Task<string> loadDepartmentPage(string link)
        {
            string page = "";

            WebClient client = new WebClient
            {
                Encoding = Encoding.GetEncoding(1251)
            };
            //client.Headers.Add("user-agent", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
            Stream data = await client.OpenReadTaskAsync(link);//сайт скачивания расписания

            StreamReader reader = new StreamReader(data, Encoding.GetEncoding(1251));
            string line = await reader.ReadLineAsync();
            while (line != null)
            {
                page += line + "\n";
                line = await reader.ReadLineAsync();
            }

            data.Close();
            reader.Close();

            return page;
        }

        public async Task<List<String>> loadFacultiesPages(string link1, string link2 = "")
        {
            string page1 = "";
            string page2 = "";

            WebClient client = new WebClient
            {
                Encoding = Encoding.GetEncoding(1251)
            };
            //client.Headers.Add("user-agent", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
            Stream data = await client.OpenReadTaskAsync(link1);//сайт скачивания расписания

            StreamReader reader = new StreamReader(data, Encoding.GetEncoding(1251));
            string line = await reader.ReadLineAsync();
            while (line != null)
            {
                page1 += line + "\n";
                line = await reader.ReadLineAsync();
            }
            data.Close();
            reader.Close();

            if (link2.Length < 1)
            {
                return new List<string>() { page1 };
            }

            data = await client.OpenReadTaskAsync(link1);//сайт скачивания расписания

            reader = new StreamReader(data, Encoding.GetEncoding(1251));
            line = await reader.ReadLineAsync();
            while (line != null)
            {
                page2 += line + "\n";
                line = await reader.ReadLineAsync();
            }
            data.Close();
            reader.Close();

            return new List<string>() { page1, page2 };// [codec.decode(pageText1), codec.decode(pageText2)];
        }

        public async Task<List<string>> loadClassroomsList()
        {
            string path = "./classrooms.txt";

            StreamReader sr = new StreamReader(path);
            var list = (await sr.ReadToEndAsync()).Split('\n').ToList();
            list.RemoveAll(match => match.Trim().Length < 1);
            for(int i = 0; i < list.Count; i++)
            {
                list[i] = Regex.Replace(list[i], "\\s", "");
            }

            return list;
        }
    }
}
