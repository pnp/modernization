using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Functions
{
    public class NotAvailableAtTargetException: Exception
    {
        public NotAvailableAtTargetException(string message): base(message) { }
        public NotAvailableAtTargetException(string message, Exception innerException) : base(message, innerException) { }
    }
}
