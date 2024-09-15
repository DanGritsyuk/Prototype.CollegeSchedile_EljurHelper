using LessonLinker.BLL.Logic;
using LessonLinker.BLL.Logic.Services;
using LessonLinker.Common.Entities.AuthEntities;
using LessonLinker.Common.Entities.ModelResponse.ApiScheduleResponse;
using LessonLinker.DAL.Repository;

namespace LessonLinker.UI
{
    public class Startup
    {
        private readonly AuthLogic _authLogic;
        private readonly GetScheduleService _scheduleService;

        private AuthData _authData => AuthData.Instance;
        private GroupListRepository _groups => new();

        private Queue<ApiScheduleResponse> _apiScheduleResponses;

        public Startup(AuthLogic authLogic)
        {
            _authLogic = authLogic;
            _scheduleService = new GetScheduleService(_groups.Groups);
            _apiScheduleResponses = new Queue<ApiScheduleResponse>();
        }

        public async Task ReadData(AuthDataRepository storage)
        {
            try
            {
                storage.LoadAuthData();
                Console.WriteLine($"Authentication data loaded...");
            }
            catch (FileNotFoundException)
            {
                await AskData();
                storage.SaveAuthData(_authData);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public async Task Start()
        {            
            var creator = new DocumentProcessor();
            creator.CreateAndMergeDocuments(@"шаблон.docx", @"E:\Расписание.docx", GenerateDocument());

            while (true)
            {

            }
        }

        private async Task AskData()
        {
            _authData.DevKey = "devkey";
            _authData.Vendor = "vendorName";
            _authData.AuthLink = new Uri("https://edu.rk.gov.ru/api/auth");
            _authData.ApiLink = new Uri("https://edu.rk.gov.ru/api/getschedule");

            await Login();
        }

        private async Task Login()
        {
            string username = "login";
            string password = "password";

            await _authLogic.GetToken(username, password);
        }

        private IAsyncEnumerable<ApiScheduleResponse> GenerateDocument()
        {
            var scheduleService = new GetScheduleService(_groups.Groups);
            return scheduleService.GetScheduleFromApi();
        }
    }
}
