using System;
using System.Collections.Generic;
using System.Text;

namespace Autocorrect.Common
{
  public static  class AppConstants
    {
        public static string ApiBaseUri = "https://localhost:5001";
        public static string SyncUri = $"{ApiBaseUri}/api/SpecialWords/getallwords";
        public static string ValidateLicenseUri = $"{ApiBaseUri}/api/license/isValid/";
        public static string UpdateUtilizationUri = $"{ApiBaseUri}/api/license/setusage/";
    }
}
