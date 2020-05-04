using System;
using System.Threading.Tasks;
using Dropbox.Api.Files;
using Dropbox.Api;
using System.IO;
using System.Text;

namespace CalculaImposto
{
    /// <summary>
    /// Esta classe faz upload de um arquivo para uma conta do dropbox e obtem um link para baixar.
    /// </summary>
    public class Program
    {

        static void Main(string[] args)
        {
            var task = Task.Run((Func<Task>)Program.Run);
            task.Wait();
        }

        static async Task Run()
        {
            //Esse token de acesso para a conta (saomarcosmateriais@gmail.com) via API. 
            using (var dbx = new DropboxClient("qiPNSnvudfAAAAAAAAAADVzCDmBXnSr2XhPH82Z2oOF24_QcWzy2lGWK0TAWKkG-"))
            {
                var full = await dbx.Users.GetCurrentAccountAsync();
                Console.WriteLine("{0} - {1}", full.Name.DisplayName, full.Email);
            }
        }

        //  file = @"C:\Users\barbi\source\repos\CalcularImpostos3\Codigo\CalculaImposto\bin\Debug\DiretorioTemporario\N_25191108475502000180550010000538061220538064_PB_000000263383599_49304497_procNFe.xml";
        // folder = @"C:\Users\barbi\source\repos\CalcularImpostos3\Codigo\CalculaImposto\bin\Debug\DiretorioTemporario";
        //  filename = "N_25191108475502000180550010000538061220538064_PB_000000263383599_49304497_procNFe.xml";
        //   url = "";

        async Task Upload(DropboxClient dbx, string folder, string file, string content)
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

