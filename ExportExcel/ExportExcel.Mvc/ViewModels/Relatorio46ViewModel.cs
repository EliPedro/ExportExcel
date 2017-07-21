using System.ComponentModel.DataAnnotations;

namespace ExportExcel.Mvc.ViewModels
{
    public class Relatorio46ViewModel
    {
        public long Registro { get; set; }
        public string Tipo { get; set; }
        [Display(Name = "Data Inicio")]
        public string DataInicio { get; set; }
        [Display(Name = "Data Final")]
        public string DataFinal { get; set; }
        public string Cliente { get; set; }
        [Display(Name = "CPF/CNPJ")]
        public string CPF_CNPJ { get; set; }
        [Display(Name = "Proprietário")]
        public string Proprietario { get; set; }
        public string Status { get; set; }
        public string Origem { get; set; }
        public string Produto { get; set; }
        public string Assunto {get;set; }
        public string Motivo { get; set; }
    }
}