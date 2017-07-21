using ExportExcel.Mvc.Code;
using ExportExcel.Mvc.Relatorio;
using System.Web.Mvc;

namespace ExportExcel.Mvc.Controllers
{
    public class RelatorioController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            return View("Index", Relatorio46.ObterRelatorio46());
        }
        
        public ActionResult ExportarExcel()
        {
            var  relatorio = Relatorio46.ObterRelatorio46();

            string[] columns = {
                                 "Registro", "Tipo", "Data Inicio", "Data Final", "Cliente","CPF / CNPJ",
                                 "Proprietário", "Status", "Origem", "Produto", "Assunto", "Motivo"
                               };
 
            byte[] filecontent = ExcelExportHelper.ExportExcel(relatorio, "Relatório de Atividades", true, columns);
            return File(filecontent, ExcelExportHelper.ExcelContentType, "Relatorio_de_Atividades_IDEA.xlsx");
        }
    }
}