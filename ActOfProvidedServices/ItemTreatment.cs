using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActOfProvidedServices {
	class ItemTreatment {
		public string Doctor { get; set; }
		public string Date { get; set; }
		public List<string> Diagnoses { get; set; } = new List<string>();
		public double TreatmentCostTotal { get; set; }
		public List<ItemService> Services { get; set; } = new List<ItemService>();
	}
}
