using System;
using System.Collections.Generic;
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace LessonLinker.Common.Entities.ModelResponse.ApiScheduleResponse
{
    public class ApiScheduleResponse
    {
        public string? GroupName { get; set; }

        [JsonPropertyName("response")]
        public ScheduleResponse Response { get; set; }
    }
}


