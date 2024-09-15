using LessonLinker.BLL.Logic;
using LessonLinker.DAL.Repository;
using LessonLinker.UI;



string filePath = "authdata.enc";
string encryptionKey = "Your32CharLongEncryptionKeyHere!";

var logic = new AuthLogic();
var app = new Startup(logic);
AuthDataRepository storage = new AuthDataRepository(filePath, encryptionKey);
await app.ReadData(storage);
app.Start();