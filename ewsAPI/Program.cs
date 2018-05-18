using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Security;
using ewsAPI.Models;
using MoreLinq;

namespace ewsAPI
{

    class Program
    {
        static void Main(string[] args)
        {
            var pf = new PublicFolder();
            var kv = new List<KeyValuePair<string, string>>();
            kv.Add(new KeyValuePair<string, string>(@"\Document Library\Agencies\Bureau of Administration\Central Mail", ""));
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var f = pf.GetAllFolders(@"sd\itpr0it231", "exchange@K#", "folder.admin@state.sd.us", kv).DistinctBy(e => e.FolderPath);
            

            watch.Stop();
            var em = ((watch.ElapsedMilliseconds / 60) / 60) / 60;
            var csv = CSVWriter.ToCsv<PublicFolderModel>(",", f);
            var stat = $"time: {em.ToString()}; NumberOfItems:{f.Count()}";
            stat.WriteFile($"D:\\PublicFolderDetails-Time_{DateTime.Now.Month}_{DateTime.Now.Day}_{DateTime.Now.Millisecond}.txt");
            csv.WriteFile($@"D:\PublicFolderDetails-Subject-From_{DateTime.Now.Month}_{DateTime.Now.Day}_{DateTime.Now.Millisecond}.csv");

            // csv.WriteFile($@"M:\_ShortTermUseOnly\ByerK\PublicFolderDetails-Subject-From_{DateTime.Now.Month}_{DateTime.Now.Month}_{DateTime.Now.Millisecond}.csv");
        }


    }


}
