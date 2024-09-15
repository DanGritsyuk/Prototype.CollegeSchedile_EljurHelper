using System.Text.Json;
using System.Text;
using LessonLinker.Common.Entities.AuthEntities;
using LessonLinker.Common.Entities.ModelResponse.ApiAuthResponse;

namespace LessonLinker.BLL.Logic
{
    public class AuthLogic
    {
        private readonly HttpClient _client;
        private readonly AuthData _authData = AuthData.Instance;

        public AuthLogic()
        {
            _client = new HttpClient();
        }

        public async Task GetToken(string username, string password)
        {
            try
            {
                await GetAuthToken(username, password);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
            }
        }

        private async Task GetAuthToken(string username, string password)
        {
            var loginData = new
            {
                login = username,
                password,
                devkey = _authData.DevKey,
                vendor = _authData.Vendor,
                out_format = "json"
            };

            var json = JsonSerializer.Serialize(loginData);
            Console.WriteLine(json);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _client.PostAsync(_authData.AuthLink, content);
            Console.WriteLine(response);
            if (response.IsSuccessStatusCode)
            {
                var jsonResponse = await response.Content.ReadAsStringAsync();
                Console.WriteLine(jsonResponse);
                var authResponse = JsonSerializer.Deserialize<ApiAuthResponse>(jsonResponse);

                _authData.Token = authResponse!.Response.Result.Token;
                _authData.SetEndDateFromString(authResponse!.Response.Result.Expires);
            }
            else
            {
                throw new Exception("Ошибка при получении токена: " + response.ReasonPhrase);
            }
        }
    }
}
