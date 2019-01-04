using Autocorrect.Api.Models;
using Autocorrect.Common;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Autocorrect.Api.Services
{
    public static class DataProvider
    {
        static HttpClient _client;
        public static Dictionary<string, string> Data;
        static DataProvider()
        {
            CreateDictionaryIfNotExists();
            _client = new HttpClient();
            System.Net.ServicePointManager.SecurityProtocol = System.Net.ServicePointManager.SecurityProtocol | System.Net.SecurityProtocolType.Tls12;

            Data = GetData();

        }
        public static string StorageFilePath{get{ return Path.Combine(StorageFolderPath,"Dictionary.json"); } }
        public static string StorageFolderPath{get{ return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ShkruajShqip"); } }
        public static Dictionary<string,string> GetData()
        {

            // deserialize JSON directly from a file
            using (StreamReader file = new StreamReader(StorageFilePath))
            {
                JsonSerializer serializer = new JsonSerializer();
                var result = (IEnumerable<WordDictionaryModel>)serializer.Deserialize(file, typeof(IEnumerable<WordDictionaryModel>));
                if (result == null) return new Dictionary<string, string>();
                return result.ToDictionary(x => x.WrongWord, y => y.RightWord, StringComparer.InvariantCultureIgnoreCase);
            }      
            
        }
        private static async Task SetData(string content)
        {
            // deserialize JSON directly from a file
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(StorageFilePath, false))
            {
                await file.WriteAsync(content);
            }
        }

        public static async Task SyncData()
        {
            var request = await _client.GetAsync(AppConstants.SyncUri);
            var content = await request.Content.ReadAsStringAsync();
            await SetData(content);
            Data = GetData();
        }
        public static void CreateDictionaryIfNotExists()
        {
            if (!Directory.Exists(StorageFolderPath)) Directory.CreateDirectory(StorageFolderPath);
            if (!File.Exists(StorageFilePath)) 
            {

                using (var tw = new StreamWriter(StorageFilePath,false))
                {
                    tw.Write("[]");
                }
            }
        }
    }
}
