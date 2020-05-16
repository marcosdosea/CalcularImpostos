using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Dropbox.Api;
using Dropbox.Api.Files;

namespace DropboxApi
{
    public class DropboxApi
    {
        public static void Main(string[] args)
        {
            var task = Task.Run((Func<Task>)DropboxApi.Run);
            task.Wait();
        }
        /// <summary>
        /// Associei o token gerado no dropbox api para o endereço de e-mail "saomarcosmateriais@gmail.com"
        /// </summary>
        /// <returns></returns>
        public static async Task Run()
        {
            using (var dbx = new DropboxClient("qiPNSnvudfAAAAAAAAAAEmV5ag8XoYVezhOQItCeNjzqJVbQDI6_M7YPa_sRWSZU"))
            {
                var full = await dbx.Users.GetCurrentAccountAsync();
                Console.WriteLine("{0} - {1}", full.Name.DisplayName, full.Email);
            }
        }
        /// <summary>
        /// Para fazer uppload de um arquivo para o dropbox
        /// </summary>
        /// <param name="dbx"></param>
        /// <param name="folder">Pasta que será criada no dropbox</param>
        /// <param name="file">Nome do arquivo com a extensão</param>
        /// <param name="content">Pasta onde se encontra o arquivo mais o nome do arquivo com extensão. Endereço completo até o arquivo</param>
        /// <returns></returns>
        public async Task Upload(DropboxClient dbx, string folder, string file, string content)
        {
            using (var mem = new MemoryStream(Encoding.UTF8.GetBytes(content)))
            {
                var updated = await dbx.Files.UploadAsync(
                    folder + "/" + file,
                    WriteMode.Overwrite.Instance,
                    body: mem);
                Console.WriteLine("Saved {0}/{1} rev {2}", folder, file, updated.Rev);
            }
        }
    }
}