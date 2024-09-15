using LessonLinker.Common.Entities.AuthEntities;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace LessonLinker.DAL.Repository
{
    public class AuthDataRepository
    {
        private readonly string _filePath;
        private readonly byte[] _encryptionKey;

        public AuthDataRepository(string filePath, string key)
        {
            _filePath = filePath;

            // Преобразуем строку ключа в байтовый массив
            _encryptionKey = Encoding.UTF8.GetBytes(key);

            if (_encryptionKey.Length != 32)
            {
                throw new ArgumentException("Ключ шифрования должен быть длиной 32 байта.");
            }
        }

        // Метод для сохранения зашифрованных данных
        public void SaveAuthData(AuthData authData)
        {
            string jsonData = JsonSerializer.Serialize(authData);
            string encryptedData = Encrypt(jsonData);

            File.WriteAllText(_filePath, encryptedData);
        }

        // Метод для загрузки и расшифровки данных
        public void LoadAuthData()
        {
            if (!File.Exists(_filePath))
            {
                throw new FileNotFoundException("Файл с данными не найден.");
            }

            string encryptedData = File.ReadAllText(_filePath);
            string decryptedData = Decrypt(encryptedData);

            // Десериализация в анонимный тип для получения полей
            var tempData = JsonSerializer.Deserialize<AuthDataTemp>(decryptedData);

            if (tempData != null)
            {
                // Заполняем текущий синглтон объект
                var authData = AuthData.Instance;
                authData.Vendor = tempData.Vendor;
                authData.Token = tempData.Token;
                authData.DevKey = tempData.DevKey;
                authData.EndDateForToken = tempData.EndDateForToken;
                authData.AuthLink = tempData.AuthLink;
                authData.ApiLink = tempData.ApiLink;
            }
        }

        // Метод шифрования
        private string Encrypt(string plainText)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = _encryptionKey;
                aes.GenerateIV();

                using (var encryptor = aes.CreateEncryptor(aes.Key, aes.IV))
                using (var memoryStream = new MemoryStream())
                {
                    memoryStream.Write(aes.IV, 0, aes.IV.Length); // Сначала записываем IV
                    using (var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                    using (var streamWriter = new StreamWriter(cryptoStream))
                    {
                        streamWriter.Write(plainText);
                    }

                    return Convert.ToBase64String(memoryStream.ToArray());
                }
            }
        }

        // Метод расшифровки
        private string Decrypt(string cipherText)
        {
            byte[] cipherBytes = Convert.FromBase64String(cipherText);

            using (Aes aes = Aes.Create())
            {
                aes.Key = _encryptionKey;
                byte[] iv = new byte[aes.BlockSize / 8];
                Array.Copy(cipherBytes, 0, iv, 0, iv.Length);
                aes.IV = iv;

                using (var decryptor = aes.CreateDecryptor(aes.Key, aes.IV))
                using (var memoryStream = new MemoryStream(cipherBytes, iv.Length, cipherBytes.Length - iv.Length))
                using (var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read))
                using (var streamReader = new StreamReader(cryptoStream))
                {
                    return streamReader.ReadToEnd();
                }
            }
        }
    }
}
