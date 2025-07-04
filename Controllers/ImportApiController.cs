using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Hosting;
using System.Web.Http;
using DemoMVCProject1.Models;
using DemoMVCProject1.Services;

namespace DemoMVCProject1.Controllers
{
    public class ImportApiController : ApiController
    {
        [HttpPost]
        [Route("api/import/importdata")]
        public IHttpActionResult ImportData([FromBody] MappingRequest request)
        {
            string path = Path.Combine(HostingEnvironment.MapPath("~/App_Data/Uploads"), Path.GetFileName(request.FileName));
            if (string.IsNullOrEmpty(path))
                return BadRequest("File path is missing.");

            var customers = ExcelService.ParseExcel(path, request.Mappings);
            ImportService.SaveCustomers(customers);
            return Ok();
        }
    }

    public class MappingRequest
    {
        public List<ExcelColumnMapping> Mappings { get; set; }
        public string FileName { get; set; }  // Add this
    }

}
