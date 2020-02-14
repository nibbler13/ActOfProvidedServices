using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActOfProvidedServices {
	class ItemService {
		public string Code { get; set; } = string.Empty;
		public string Name { get; set; } = string.Empty;
		public double Count { get; set; } = 0;
		public double Cost { get; set; } = 0;
		public double CostFinal { get; set; } = 0;
		public string ToothNumber { get; set; } = string.Empty;
	}
}
