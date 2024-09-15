using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace LessonLinker.Common.Entities.ModelResponse.ApiScheduleResponse
{
    public class LessonItem
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("num")]
        public string Number { get; set; }

        [JsonPropertyName("room")]
        public string Room { get; set; }

        [JsonPropertyName("teacher")]
        public string Teacher { get; set; }

        [JsonPropertyName("sort")]
        public object Sort { get; set; }

        [JsonPropertyName("teacher_id")]
        public int TeacherId { get; set; }

        [JsonPropertyName("grp_short")]
        public string GroupShort { get; set; }

        [JsonPropertyName("grp")]
        public string Group { get; set; }
    }
}
