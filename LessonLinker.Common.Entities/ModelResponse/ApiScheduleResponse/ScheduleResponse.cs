using System.Text.Json.Serialization;

namespace LessonLinker.Common.Entities.ModelResponse.ApiScheduleResponse
{
    public class ScheduleResponse
    {
        [JsonPropertyName("state")]
        public int State { get; set; }

        [JsonPropertyName("error")]
        public string Error { get; set; }

        [JsonPropertyName("result")]
        public ScheduleResult Result { get; set; }
    }
}
