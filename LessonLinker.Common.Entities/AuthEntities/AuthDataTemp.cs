using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LessonLinker.Common.Entities.AuthEntities
{
    public class AuthDataTemp
    {
        public string Vendor { get; set; }
        public string Token { get; set; }
        public string DevKey { get; set; }
        public DateTime EndDateForToken { get; set; }
        public Uri AuthLink { get; set; }
        public Uri ApiLink { get; set; }
    }
}