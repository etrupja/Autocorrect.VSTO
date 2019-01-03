using Autocorrect.Api.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Autocorrect.Api.Services
{
    public class DataProvider
    {
        public Dictionary<string,string> GetData()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Data/Dictionary.json");
            // deserialize JSON directly from a file
            using (StreamReader file = new StreamReader(filePath, Encoding.GetEncoding("iso-8859-1")))
            {
                JsonSerializer serializer = new JsonSerializer();
                var result = (IEnumerable<WordDictionaryModel>)serializer.Deserialize(file, typeof(IEnumerable<WordDictionaryModel>));
                return result.ToDictionary(x => x.Wrong, y => y.Right, StringComparer.InvariantCultureIgnoreCase);
            }
        
            
        }
    }
}
