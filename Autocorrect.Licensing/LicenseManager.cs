using Autocorrect.Common;
using Newtonsoft.Json;
using Portable.Licensing;
using Portable.Licensing.Validation;
using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace Autocorrect.Licensing
{
    public static class LicenseManager
    {
        public const string LicensePublicKey = "MIIBKjCB4wYHKoZIzj0CATCB1wIBATAsBgcqhkjOPQEBAiEA/////wAAAAEAAAAAAAAAAAAAAAD///////////////8wWwQg/////wAAAAEAAAAAAAAAAAAAAAD///////////////wEIFrGNdiqOpPns+u9VXaYhrxlHQawzFOw9jvOPD4n0mBLAxUAxJ02CIbnBJNqZnjhE50mt4GffpAEIQNrF9Hy4SxCR/i85uVjpEDydwN9gS3rM6D0oTlF2JjClgIhAP////8AAAAA//////////+85vqtpxeehPO5ysL8YyVRAgEBA0IABPCLaFbzw/MJhWd/DzjPNKSgd9/fz6Jo0oSJHt3PTNNGLzzppCZwuJ8Mwvkw0ARHYgCfzIxiXKfSedSDyRBO5lo=";

        static HttpClient _client;
        
        static LicenseManager()
        {
            if (!Directory.Exists(LicenseFolderPath)) Directory.CreateDirectory(LicenseFolderPath);
            System.Net.ServicePointManager.SecurityProtocol = System.Net.ServicePointManager.SecurityProtocol | System.Net.SecurityProtocolType.Tls12;

            _client = new HttpClient();
            License = GetLicense();
        }
        public static License License { get; private set; }
        private static string LicenseFilePath { get { return Path.Combine(LicenseFolderPath, "License.xml"); } }
        private static string LicenseFolderPath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ShkruajShqip"); } }


        private static License GetLicense()
        {
            if (!File.Exists(LicenseFilePath)) return null;
            using (StreamReader file = new StreamReader(LicenseFilePath))
            {
                try
                {
                    return License.Load(file);
                }
                catch
                {

                    return null;
                }
            }
        }
        public static bool IsLicenseValid()
        {
            if (License == null) return false;
            return IsValid(License);
        }
        public static bool IsValid(License license)
        {
            var validationFailures = license.Validate()
                                .ExpirationDate()
                                .And()
                                .Signature(LicensePublicKey).AssertValidLicense();
            return !validationFailures.Any();
        }
        public static DateTime? ExpirationDate()
        {
            if (License == null) return null;
            return License.Expiration;

        }
        public static async Task<bool> ValidateLicenseOnline(Guid id)
        {
            var request = await _client.GetAsync(AppConstants.ValidateLicenseUri + id.ToString());
            var content = await request.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<bool>(content);
        }
        public static async Task UpdateLicenseUtilizedCount(Guid id)
        {
            var request = await _client.PostAsync(AppConstants.UpdateUtilizationUri + id.ToString(),null);
            request.EnsureSuccessStatusCode();
        }
        public static async Task SetLicense(Stream data)
        {
            data.Position = 0;
            var fileStream = new FileStream(LicenseFilePath, FileMode.OpenOrCreate, FileAccess.Write);
            await data.CopyToAsync(fileStream);
            fileStream.Dispose();
            License = GetLicense();
        }
        public static bool HasLicense()
        {
            return License != null;
        }

        public static License ParseLicense(Stream data)
        {
            return License.Load(data);
        }


    }
}
