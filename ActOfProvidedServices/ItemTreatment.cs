using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActOfProvidedServices {
	class ItemTreatment {
		public string Doctor { get; set; } = string.Empty;
		public string Date { get; set; } = string.Empty;
		public List<string> Diagnoses { get; set; } = new List<string>();
		public double TreatmentCostTotal { get; set; } = 0;
		public string Filial { get; set; } = string.Empty;
		public string GarantyMail { get; set; } = string.Empty;
		public List<ItemService> Services { get; set; } = new List<ItemService>();
	}
}
