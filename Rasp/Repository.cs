using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SheduleSI
{
    class Repository
    {
        readonly HttpClient client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(3),            
        };

        public async Task<string> LoadDepartmentPage(string link)
        {
            var response = await client.GetByteArrayAsync(link);
            var page = Encoding.GetEncoding(1251).GetString(response, 0, response.Length);
            
            return page;
        }

        public async Task<string[]> LoadFacultiesPages(string link1, string link2 = "")
        {
            var response = await client.GetByteArrayAsync(link1);
            var page1 = Encoding.GetEncoding(1251).GetString(response, 0, response.Length);

            if (link2.Length < 1)
            {
                return new string[] { page1 };
            }

            var response2 = await client.GetByteArrayAsync(link2);
            var page2 = Encoding.GetEncoding(1251).GetString(response2, 0, response2.Length);

            return new string[] { page1, page2 };
        }

        public async Task<List<string>> LoadClassroomsList()
        {
            string path = "./classrooms.txt";

            StreamReader sr = new StreamReader(path);
            var list = (await sr.ReadToEndAsync()).Split('\n').ToList();
            sr.Close();
            list.RemoveAll(match => match.Trim().Length < 1);
            for(int i = 0; i < list.Count; i++)
            {
                list[i] = Regex.Replace(list[i], "\\s", "");
            }

            return list;
        }
    }
}
