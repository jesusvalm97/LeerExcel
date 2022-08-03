using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ApiLeerExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LeerExcelController : ControllerBase
    {
        // GET: api/<LeerExcelController>
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/<LeerExcelController>/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }

        // POST api/<LeerExcelController>
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/<LeerExcelController>/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/<LeerExcelController>/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
