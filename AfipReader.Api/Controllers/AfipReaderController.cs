using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Aspose;
using AfipReader.Core;
using Aspose.Cells;
// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace AfipReader.Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AfipReaderController : ControllerBase
    {
        /*
        // GET: api/<AfipReaderController>
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/<AfipReaderController>/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }
        */
        // POST api/<AfipReaderController>
        [HttpPost]
        public (IEnumerable<Comprobante>,IEnumerable<int>) Post()
        {
            if (Request.Form.Files.Count == 0)
            {
                throw new InvalidOperationException(" No Files");
            }

            IFormFile file = Request.Form.Files[0];
            var stream = file.OpenReadStream();


            Workbook workbook = new Workbook (stream);
            
            AfipWorksheet sheet = new AfipWorksheet();
            return (sheet.GetDetails(workbook).Item1, sheet.GetDetails(workbook).Item2);

        }
        /*
        // PUT api/<AfipReaderController>/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/<AfipReaderController>/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
        */
    }
}
