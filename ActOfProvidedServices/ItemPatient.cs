using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActOfProvidedServices {
	class ItemPatient {
		public string Name { get; set; } = string.Empty;
		public string Documents { get; set; } = string.Empty;
		public string Code { get; set; } = string.Empty;
		public string GarantyMail { get; set; } = string.Empty;
		public double PatientCostTotal { get; set; } = 0;
		public List<ItemTreatment> Treatments { get; set; } = new List<ItemTreatment>();
	}
}
