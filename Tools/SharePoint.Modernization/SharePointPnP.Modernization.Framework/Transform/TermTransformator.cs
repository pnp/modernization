using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Transform
{
    public  class TermTransformator : BaseTransform
    {
        private ClientContext _sourceContext;
        private ClientContext _targetContext;

        #region Construction        

        /// <summary>
        /// Constructor for the Term Transformator class
        /// </summary>
        /// <param name="baseTransformationInformation"></param>
        /// <param name="sourceContext"></param>
        /// <param name="targetContext"></param>
        /// <param name="logObservers"></param>
        public TermTransformator(BaseTransformationInformation baseTransformationInformation, ClientContext sourceContext, ClientContext targetContext, IList<ILogObserver> logObservers = null)
        {
            // Hookup logging
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            // Ensure source and target context are set
            if (sourceContext == null && targetContext != null)
            {
                sourceContext = targetContext;
            }

            if (targetContext == null && sourceContext != null)
            {
                targetContext = sourceContext;
            }


            this._sourceContext = sourceContext;
            this._targetContext = targetContext;

        }

        #endregion


    }
}
