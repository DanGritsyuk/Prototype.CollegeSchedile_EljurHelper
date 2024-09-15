using System.Globalization;

namespace LessonLinker.Common.Entities.AuthEntities
{
    public class AuthData
    {
        private static readonly AuthData _instance = new AuthData();

        private AuthData() { }

        public static AuthData Instance
        {
            get { return _instance; }
        }

        public string Vendor { get; set; }
        public string Token { get; set; }
        public string DevKey { get; set; }
        public DateTime EndDateForToken { get; set; }
        public Uri AuthLink { get; set; }
        public Uri ApiLink { get; set; }


        public void Initialize(string vendor, string token, string devKey, string endDateForToken, Uri authLink, Uri apiLink)
        {
            Vendor = vendor;
            Token = token;
            DevKey = devKey;
            EndDateForToken = ParseDate(endDateForToken);
            AuthLink = authLink;
            AuthLink = authLink;
        }

        public void SetEndDateFromString(string str) =>
            EndDateForToken = ParseDate(str);

        private static DateTime ParseDate(string dateString)
        {
            string format = "yyyy-MM-dd HH:mm:ss";
            DateTime dateTime;

            try
            {
                dateTime = DateTime.ParseExact(dateString, format, CultureInfo.InvariantCulture);
                return dateTime;
            }
            catch (FormatException)
            {
                throw;
            }
        }
    }
}
