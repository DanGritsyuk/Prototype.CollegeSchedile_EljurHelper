using LessonLinker.Common.Entities.AuthEntities;
using LessonLinker.Common.Entities.ModelResponse.ApiScheduleResponse;
using System.Text;
using System.Text.Json;

namespace LessonLinker.BLL.Logic.Services
{
    public class GetScheduleService
    {
        private readonly HttpClient _client;

        private readonly IEnumerable<string> _groups;
        private AuthData _authData => AuthData.Instance;

        public GetScheduleService(IEnumerable<string> groups)
        {
            _client = new HttpClient();
            _groups = groups;
        }

        public async IAsyncEnumerable<ApiScheduleResponse> GetScheduleFromApi()
        {
            foreach (var group in _groups)
            {
                yield return await GetSchedule(group);
            }
        }

        private async Task<ApiScheduleResponse> GetSchedule(string group)
        {
            // Формируем строку запроса с параметрами
            var queryString = $"?devkey={Uri.EscapeDataString(_authData.DevKey)}" +
                              $"&vendor={Uri.EscapeDataString(_authData.Vendor)}" +
                              $"&auth_token={Uri.EscapeDataString(_authData.Token)}" +
                              $"&class={Uri.EscapeDataString(group)}" +
                              $"&days=20241007-20241013" + // data for tests 20240930-20241006
                              $"&out_format=json";

            // Формируем полный URL
            var requestUrl = _authData.ApiLink + queryString;

            // Выполняем GET-запрос
            var response = await _client.GetAsync(requestUrl);

            if (response.IsSuccessStatusCode)
            {
                var jsonResponse = await response.Content.ReadAsStringAsync();
                var result = JsonSerializer.Deserialize<ApiScheduleResponse>(jsonResponse)!;
                result.GroupName = group;
                return result;
            }
            else
            {
                throw new Exception($"Ошибка при получении расписания для {group}: {response.ReasonPhrase}");
            }
        }

    }
}