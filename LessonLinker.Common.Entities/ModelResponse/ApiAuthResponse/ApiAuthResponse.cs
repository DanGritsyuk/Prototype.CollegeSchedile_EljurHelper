using System;
using System.Collections.Generic;
using System.Formats.Asn1;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace LessonLinker.Common.Entities.ModelResponse.ApiAuthResponse
{
    public class ApiAuthResponse
    {
        [JsonPropertyName("response")]
        public AuthResponse Response { get; set; }
    }
}
