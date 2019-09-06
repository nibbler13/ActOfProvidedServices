using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActOfProvidedServices {
	class ItemPatient {
		public string Name { get; set; }
		public string Documents { get; set; }
		public string Code { get; set; }
		public double PatientCostTotal { get; set; }
		public List<ItemTreatment> Treatments { get; set; } = new List<ItemTreatment>();
	}
}
