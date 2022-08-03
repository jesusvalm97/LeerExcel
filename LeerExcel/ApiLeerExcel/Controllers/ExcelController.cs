using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ApiLeerExcel.Models;

namespace ApiLeerExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpGet]
        public ActionResult<Models.ObjetoRespuesta> GetResult()
        {
            Models.ObjetoRespuesta objetoRespuesta = new Models.ObjetoRespuesta();
            objetoRespuesta.Propiedades.Add("Saludos", "Hola mundo");

            return Ok(objetoRespuesta);
        }

        [HttpPost]
        public ActionResult<Models.ObjetoRespuesta> PostResult(ObjetoRespuesta objetoRespuesta)
        {
            return Ok(objetoRespuesta);
        }
    }
}
