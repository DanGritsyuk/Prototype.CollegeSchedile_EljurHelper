using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace LessonLinker.Common.Entities.ModelResponse.ApiScheduleResponse
{
    public class Day
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("title")]
        public string Title { get; set; }

        [JsonPropertyName("alert")]
        public string Alert { get; set; }

        [JsonPropertyName("items")]
        public List<LessonItem> Items { get; set; }
    }
}
