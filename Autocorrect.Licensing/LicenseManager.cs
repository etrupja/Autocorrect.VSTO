using Portable.Licensing;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Autocorrect.Licensing
{
    public class LicenseManager
    {
        public const string LicensePublicKey = "MIIBKjCB4wYHKoZIzj0CATCB1wIBATAsBgcqhkjOPQEBAiEA/////wAAAAEAAAAAAAAAAAAAAAD///////////////8wWwQg/////wAAAAEAAAAAAAAAAAAAAAD///////////////wEIFrGNdiqOpPns+u9VXaYhrxlHQawzFOw9jvOPD4n0mBLAxUAxJ02CIbnBJNqZnjhE50mt4GffpAEIQNrF9Hy4SxCR/i85uVjpEDydwN9gS3rM6D0oTlF2JjClgIhAP////8AAAAA//////////+85vqtpxeehPO5ysL8YyVRAgEBA0IABPCLaFbzw/MJhWd/DzjPNKSgd9/fz6Jo0oSJHt3PTNNGLzzppCZwuJ8Mwvkw0ARHYgCfzIxiXKfSedSDyRBO5lo=";
        public LicenseManager()
        {
            if (!Directory.Exists(LicenseFolderPath)) Directory.CreateDirectory(LicenseFolderPath);
        }
        private  string LicenseFilePath { get { return Path.Combine(LicenseFolderPath, "License.xml"); } }
        private  string LicenseFolderPath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ShkruajShqip"); } }
        private License GetLicense()
        {
            if (!File.Exists(LicenseFilePath)) return null;
            using (StreamReader file = new StreamReader(LicenseFilePath))
            {
                return License.Load(file);
            }
        }
        public bool IsLicenseValid()
        {
            var license = GetLicense();
            if (license == null) return false;
            return license.VerifySignature(LicensePublicKey);
        }

        public DateTime? ExpirationDate()
        {
            var license = GetLicense();
            if (license == null) return null;
            return license.Expiration;

        }

        public async Task SetLicense(Stream data)
        {
            var fileStream = new FileStream(LicenseFilePath, FileMode.OpenOrCreate, FileAccess.Write);
            await data.CopyToAsync(fileStream);
            fileStream.Dispose();
        }
        public bool HasLicense()
        {
            return GetLicense() != null;
        }

    }
}
