using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

using Alma.NetKernel.Tools;
using Alma.BaseUI.ErrorMessage;

using Wpm.Implement.Manager;
using Wpm.Schema.Kernel;

namespace Actcut.ActcutClipperApi
{
    public interface IClipperApi
    {
        bool Init(IContext Context);
        bool Init(string DatabaseName, string User);

        bool ExportQuote(long QuoteNumber, string OrderNumber, string ExportFile);
        bool GetQuote(out long QuoteNumberReference);
    }
}
