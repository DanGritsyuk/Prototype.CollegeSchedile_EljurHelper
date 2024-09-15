using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace LessonLinker.Common.Entities.ModelResponse.ApiAuthResponse
{
    public class AuthResult
    {
        [JsonPropertyName("token")]
        public string Token { get; set; }

        [JsonPropertyName("expires")]
        public string Expires { get; set; }
    }
}
