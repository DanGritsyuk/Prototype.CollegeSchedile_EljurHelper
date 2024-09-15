using System.Text.Json.Serialization;

namespace LessonLinker.Common.Entities.ModelResponse.ApiAuthResponse
{
    public class AuthResponse
    {
        [JsonPropertyName("state")]
        public int State { get; set; }
        [JsonPropertyName("error")]
        public object Error { get; set; }
        [JsonPropertyName("result")]
        public AuthResult Result { get; set; }
    }
}
