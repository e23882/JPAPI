using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace JPAPI.Controllers
{
    
    [Route("api/JPAPI")]
    //[ApiController]
    public class JPController : ControllerBase
    {
        /// <summary>
        /// 取得指定日期工作時數
        /// </summary>
        /// <param name="id"></param>
        /// <param name="pw"></param>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <param name="day"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetWorkTime")]
        public ActionResult<string> GetWorkTime(string id, string pw, string year, string month, string day)
        {
            string resultText = string.Empty;
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = @"C:\Users\LeoYang\AppData\Local\Programs\Python\Python38-32\python.exe";
            start.Arguments = string.Format("{0} {1}", @"C:\Users\LeoYang\source\repos\JPAPI\JPAPI\bin\Debug\netcoreapp2.1\SearchWorkTime.py", $"{id} {pw} {year} {month} {day}"); 
            start.UseShellExecute = false;
            start.RedirectStandardOutput = true;
            using (Process process = Process.Start(start))
            {
                using (StreamReader reader = process.StandardOutput)
                {
                    string result = reader.ReadToEnd();
                    resultText += result;
                }
            }
            return resultText;
        }

        /// <summary>
        /// 計算加班
        /// </summary>
        /// <param name="id"></param>
        /// <param name="pw"></param>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <param name="day"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("CalculateOverWork")]
        public ActionResult<string> CalculateOverWork(string id, string pw, string year, string month)
        {
            string resultText = string.Empty;
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = @"C:\Users\LeoYang\AppData\Local\Programs\Python\Python38-32\python.exe";
            start.Arguments = string.Format("{0} {1}", @"C:\Users\LeoYang\source\repos\JPAPI\JPAPI\bin\Debug\netcoreapp2.1\CalculateOvertime.py", $"{id} {pw} {year} {month}");
            start.UseShellExecute = false;
            start.RedirectStandardOutput = true;
            using (Process process = Process.Start(start))
            {
                using (StreamReader reader = process.StandardOutput)
                {
                    string result = reader.ReadToEnd();
                    resultText += result;
                }
            }
            return resultText;
        }

        /// <summary>
        /// 登入工時
        /// </summary>
        /// <param name="id"></param>
        /// <param name="pw"></param>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <param name="day"></param>
        /// <param name="workTime"></param>
        /// <param name="projectCode"></param>
        /// <param name="workType"></param>
        /// <param name="team"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("LoginWorkTime")]
        public ActionResult<string> LoginWorkTime(string id, string pw, string year, string month, string day, int workTime, string projectCode, string workType, int team)
        {
            string resultText = string.Empty;
            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = @"C:\Users\LeoYang\AppData\Local\Programs\Python\Python38-32\python.exe";
            start.Arguments = string.Format("{0} {1}", @"C:\Users\LeoYang\source\repos\JPAPI\JPAPI\bin\Debug\netcoreapp2.1\LoginWorktime.py", $"{id} {pw} {year} {month} {day} {workTime} {projectCode} {workType} {team}");
            start.UseShellExecute = false;
            start.RedirectStandardOutput = true;
            using (Process process = Process.Start(start))
            {
                using (StreamReader reader = process.StandardOutput)
                {
                    string result = reader.ReadToEnd();
                    resultText += result;
                }
            }
            return resultText;
        }

        //[HttpGet]
        //[Route("AskLeave")]
        //public ActionResult<string> AskLeave(string id, string pw, string year, string month, string day, string startHR, string endHR)
        //{
        //    string resultText = string.Empty;
        //    ProcessStartInfo start = new ProcessStartInfo();
        //    start.FileName = @"C:\Users\LeoYang\AppData\Local\Programs\Python\Python38-32\python.exe";
        //    start.Arguments = string.Format("{0} {1}", @"C:\Users\LeoYang\source\repos\JPAPI\JPAPI\bin\Debug\netcoreapp2.1\AskLeave.py", $"{id} {pw} {year} {month} {day} {startHR} {endHR}");
        //    start.UseShellExecute = false;
        //    start.RedirectStandardOutput = true;
        //    using (Process process = Process.Start(start))
        //    {
        //        using (StreamReader reader = process.StandardOutput)
        //        {
        //            string result = reader.ReadToEnd();
        //            resultText += result;
        //        }
        //    }
        //    return resultText;
        //}
    }
}