using System.Text.Json.Serialization;

namespace LessonLinker.Common.Entities.ModelResponse.ApiScheduleResponse
{
    public class ScheduleResult
    {
        [JsonPropertyName("days")]
        public Dictionary<string, Day> Days { get; set; }
    }
}